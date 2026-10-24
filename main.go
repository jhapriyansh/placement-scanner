package main

import (
	"bytes"
	"encoding/json"
	"io"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	imap "github.com/emersion/go-imap"
	imapclient "github.com/emersion/go-imap/client"
	"github.com/emersion/go-message/mail"
	"github.com/joho/godotenv"
	"github.com/xuri/excelize/v2"
)

type Config struct {
	ImapServer   string
	EmailUser    string
	EmailPass    string
	MyName       string
	MyNeoID      string
	MyRegNo      string
	NtfyURL      string
	NtfyUser     string
	NtfyPass     string
	CheckInterval time.Duration
	MaxFileSizeMB int64
}

const StateFile = "processed_uids.json"

func loadConfig() Config {
	_ = godotenv.Load()
	intervalSec, _ := strconv.Atoi(os.Getenv("CHECK_INTERVAL"))
	if intervalSec <= 0 {
		intervalSec = 300
	}
	maxSize, _ := strconv.ParseInt(os.Getenv("MAX_FILE_SIZE_MB"), 10, 64)
	if maxSize <= 0 {
		maxSize = 10
	}

	return Config{
		ImapServer:    os.Getenv("IMAP_SERVER"),
		EmailUser:     os.Getenv("EMAIL_USER"),
		EmailPass:     os.Getenv("EMAIL_PASS"),
		MyName:        os.Getenv("MY_NAME"),
		MyNeoID:       os.Getenv("MY_NEO_ID"),
		MyRegNo:       os.Getenv("MY_REG_NO"),
		NtfyURL:       os.Getenv("NTFY_URL"),
		NtfyUser:      os.Getenv("NTFY_USER"),
		NtfyPass:      os.Getenv("NTFY_PASS"),
		CheckInterval: time.Duration(intervalSec) * time.Second,
		MaxFileSizeMB: maxSize,
	}
}

func loadProcessedUIDs() map[uint32]bool {
	uids := make(map[uint32]bool)
	if _, err := os.Stat(StateFile); os.IsNotExist(err) {
		return uids
	}
	data, err := os.ReadFile(StateFile)
	if err != nil {
		return uids
	}
	var uidList []uint32
	if json.Unmarshal(data, &uidList) == nil {
		for _, uid := range uidList {
			uids[uid] = true
		}
	}
	return uids
}

func saveProcessedUID(uid uint32) {
	uids := loadProcessedUIDs()
	uids[uid] = true
	var list []uint32
	for k := range uids {
		list = append(list, k)
	}
	// Keep last 1000 UIDs
	if len(list) > 1000 {
		list = list[len(list)-1000:]
	}
	data, _ := json.Marshal(list)
	_ = os.WriteFile(StateFile, data, 0644)
}

func sendNtfy(cfg Config, title, message string) {
	req, err := http.NewRequest("POST", cfg.NtfyURL, strings.NewReader(message))
	if err != nil {
		log.Printf("Failed to create ntfy request: %v", err)
		return
	}
	req.SetBasicAuth(cfg.NtfyUser, cfg.NtfyPass)
	req.Header.Set("Title", title)
	req.Header.Set("Priority", "urgent")
	req.Header.Set("Tags", "bell,briefcase")

	client := &http.Client{Timeout: 10 * time.Second}
	resp, err := client.Do(req)
	if err != nil {
		log.Printf("Failed to send ntfy notification: %v", err)
		return
	}
	defer resp.Body.Close()
}

func matchesDetails(text string, cfg Config) bool {
	lower := strings.ToLower(text)
	if cfg.MyRegNo != "" && strings.Contains(lower, strings.ToLower(cfg.MyRegNo)) {
		return true
	}
	if cfg.MyNeoID != "" && strings.Contains(lower, strings.ToLower(cfg.MyNeoID)) {
		return true
	}
	if cfg.MyName != "" && strings.Contains(lower, strings.ToLower(cfg.MyName)) {
		return true
	}
	return false
}

func scanExcel(fileBytes []byte, cfg Config) bool {
	f, err := excelize.OpenReader(bytes.NewReader(fileBytes))
	if err != nil {
		log.Printf("Error opening Excel file: %v", err)
		return false
	}
	defer f.Close()

	for _, sheet := range f.GetSheetList() {
		rows, err := f.GetRows(sheet)
		if err != nil {
			continue
		}
		for _, row := range rows {
			for _, cell := range row {
				if matchesDetails(cell, cfg) {
					return true
				}
			}
		}
	}
	return false
}

func checkEmails(cfg Config) {
	log.Println("Checking inbox for placement updates...")
	c, err := imapclient.DialTLS(cfg.ImapServer, nil)
	if err != nil {
		log.Printf("IMAP connection error: %v", err)
		return
	}
	defer c.Logout()

	if err := c.Login(cfg.EmailUser, cfg.EmailPass); err != nil {
		log.Printf("IMAP login error: %v", err)
		return
	}

	_, err = c.Select("INBOX", false)
	if err != nil {
		log.Printf("IMAP select inbox error: %v", err)
		return
	}

	processed := loadProcessedUIDs()

	criteria := imap.NewSearchCriteria()
	criteria.WithoutFlags = []string{imap.DeletedFlag}
	uids, err := c.Search(criteria)
	if err != nil {
		log.Printf("IMAP search error: %v", err)
		return
	}

	if len(uids) == 0 {
		return
	}

	seqSet := new(imap.SeqSet)
	seqSet.AddNum(uids...)

	messages := make(chan *imap.Message, 10)
	done := make(chan error, 1)

	go func() {
		done <- c.Fetch(seqSet, []string{"UID", imap.FetchEnvelope, imap.FetchBody}, messages)
	}()

	for msg := range messages {
		if processed[msg.Uid] {
			continue
		}

		subject := ""
		if msg.Envelope != nil {
			subject = msg.Envelope.Subject
		}

		matched := false
		var bodyText strings.Builder

		for _, literal := range msg.Body {
			mr, err := mail.CreateReader(literal)
			if err != nil {
				continue
			}

			for {
				p, err := mr.NextPart()
				if err == io.EOF {
					break
				} else if err != nil {
					break
				}

				switch h := p.Header.(type) {
				case *mail.InlineHeader:
					contentType, _, _ := h.ContentType()
					if strings.HasPrefix(contentType, "text/plain") {
						b, _ := io.ReadAll(p.Body)
						bodyText.Write(b)
					}
				case *mail.AttachmentHeader:
					filename, _ := h.Filename()
					if filename != "" && (strings.HasSuffix(strings.ToLower(filename), ".xlsx") || strings.HasSuffix(strings.ToLower(filename), ".xls")) {
						fileBytes, _ := io.ReadAll(p.Body)
						sizeMB := int64(len(fileBytes)) / (1024 * 1024)

						if sizeMB > cfg.MaxFileSizeMB {
							log.Printf("⚠️ Large attachment skipped: %s (%dMB)", filename, sizeMB)
							sendNtfy(cfg, "⚠️ Large Attachment Skipped", 
								"Email Subject: "+subject+"\nFile: "+filename+" exceeded size limit. Please check manually!")
							continue
						}

						if scanExcel(fileBytes, cfg) {
							matched = true
						}
					}
				}
			}
		}

		if matchesDetails(bodyText.String(), cfg) || matched {
			log.Printf("Match found! Sending alert for: %s", subject)
			sendNtfy(cfg, "🎉 Shortlisted: "+subject, "Your details were matched in email or attachment!")
		}

		saveProcessedUID(msg.Uid)
	}

	if err := <-done; err != nil {
		log.Printf("Fetch error: %v", err)
	}
}

func main() {
	cfg := loadConfig()
	log.Println("Go Placement Scanner Started...")

	for {
		checkEmails()
		time.sleep(cfg.CheckInterval)
	}
}
