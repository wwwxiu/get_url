package main

import (
	"bufio"
	"context"
	"crypto/tls"
	"fmt"
	"log"
	"net"
	"net/http"
	"os"
	"os/signal"
	"path/filepath"
	"strings"
	"sync"
	"sync/atomic"
	"syscall"
	"time"

	"github.com/fatih/color"
	"github.com/urfave/cli/v3"
	"github.com/xuri/excelize/v2"
	"golang.org/x/net/html"
)

type Config struct {
	File    string
	Count   int
	Timeout int
	OutFile string
	Tls     bool
}
type UrlResult struct {
	Url    string
	Status int
	Title  string
	Error  error
}

type Stats struct {
	Url       int32
	Success   int32
	ForBidden int32
	NotFound  int32
	Timeout   int32
	Failed    int32
	TotTime   int32
}

func main() {
	sigChan := make(chan os.Signal, 1)
	signal.Notify(sigChan, syscall.SIGINT, syscall.SIGTERM)
	go func() {
		sig := <-sigChan
		fmt.Println("%V, exit", sig)
		os.Exit(0)
	}()

	cfg := CliCommand()
	urls, err := HandleFile(cfg.File)
	if err != nil {
		log.Fatal(err)
	}
	stat := &Stats{}
	fmt.Printf("开始处理%d个Url, 并发: %d\n", len(urls), cfg.Count)
	urlChan := make(chan string, cfg.Count*2)
	resultChan := make(chan UrlResult, cfg.Count*2)
	stopChan := make(chan struct{})
	doneChan := make(chan bool)
	go WriteResult(stat, resultChan, cfg, doneChan)

	var wg sync.WaitGroup
	for i := 0; i < cfg.Count; i++ {
		wg.Add(1)
		go worker(cfg, urlChan, resultChan, stat, &wg, stopChan)
	}
	go func() {
		for _, url := range urls {
			select {
			case <-stopChan:
				return
			case urlChan <- url:

			}
		}
		close(urlChan)
	}()
	wg.Wait()
	close(resultChan)
	<-doneChan
	printStats(stat)

}

func CliCommand() *Config {

	cfg := &Config{}
	cmd := &cli.Command{
		Name:  "speter",
		Usage: "Get Url",
		Flags: []cli.Flag{
			&cli.StringFlag{
				Name:        "file",
				Usage:       "Excel文件路径",
				Aliases:     []string{"f"},
				Required:    true,
				Destination: &cfg.File,
			},
			&cli.IntFlag{
				Name:        "count",
				Usage:       "并发数量",
				Aliases:     []string{"c"},
				Value:       20,
				Destination: &cfg.Count,
			},
			&cli.IntFlag{
				Name:        "timeout",
				Usage:       "超时时间",
				Aliases:     []string{"t"},
				Value:       5,
				Destination: &cfg.Timeout,
			},
			&cli.StringFlag{
				Name:        "outfile",
				Usage:       "输出Excel文件",
				Aliases:     []string{"o"},
				Destination: &cfg.OutFile,
			},
			&cli.BoolFlag{
				Name:        "tls-skip",
				Usage:       "跳过TLS证书检查",
				Value:       false,
				Destination: &cfg.Tls,
			},
		},

		Action: func(ctx context.Context, cli *cli.Command) error {
			if cfg.Timeout <= 0 {
				return fmt.Errorf("超时时间不能为0或更小")
			}
			if cfg.Count <= 0 {
				return fmt.Errorf("并发数不能为0")
			}
			return nil
		},
	}
	if err := cmd.Run(context.Background(), os.Args); err != nil {
		log.Fatal(err)
	}
	return cfg
}

func HandleFile(FilePath string) ([]string, error) {
	var urls []string
	ext := strings.ToLower(filepath.Ext(FilePath))
	if ext == ".xlsx" {
		file, err := excelize.OpenFile(FilePath)
		if err != nil {
			return nil, fmt.Errorf("无法打开文件")
		}
		defer file.Close()
		sheet := file.GetSheetList()[0]
		rows, err := file.GetRows(sheet)
		if err != nil {
			return nil, fmt.Errorf("get line error")
		}
		urls := []string{}
		for RowIndex, row := range rows {
			//跳过表头
			if RowIndex == 0 {
				continue
			}
			if len(row) > 0 && row[0] != "" {
				urls = append(urls, row[0])
			}
		}
		return urls, nil
	}
	if ext == ".txt" {
		if _, err := os.Stat(FilePath); os.IsNotExist(err) {
			return nil, fmt.Errorf("文件不存在")
		}
		file, err := os.Open(FilePath)
		if err != nil {
			return nil, fmt.Errorf("文件打开错误")
		}
		defer file.Close()
		scanner := bufio.NewScanner(file)
		for scanner.Scan() {
			line := strings.TrimSpace(scanner.Text())
			if line == "" {
				continue
			}
			urls = append(urls, line)
		}
		return urls, nil
	}

	return nil, fmt.Errorf("不支持该文件格式")

}

func fetchUrl(url string, timeout int, verify bool) (int, string, error) {

	client := &http.Client{
		Timeout: time.Duration(timeout) * time.Second,
		Transport: &http.Transport{
			TLSClientConfig: &tls.Config{
				InsecureSkipVerify: verify,
			},
		},
	}
	if !strings.HasPrefix(url, "http://") && !strings.HasPrefix(url, "https://") {
		url = "https://" + url
	}

	req, err := http.NewRequest("GET", url, nil)
	if err != nil {
		return 0, "", fmt.Errorf("建立连接失败")
	}
	req.Header.Set(
		"User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/143.0.0.0 Safari/537.36 Edg/143.0.0.0")
	resp, err := client.Do(req)
	if err != nil {
		if ne, ok := err.(net.Error); ok && ne.Timeout() {
			return -1, "", fmt.Errorf("timeout")
		}
		return 0, "", err
	}
	defer resp.Body.Close()
	doc, err := html.Parse(resp.Body)
	if err != nil {
		return resp.StatusCode, "网页没有内容", nil
	}
	if title := findTitle(doc); title != "" {
		return resp.StatusCode, title, nil
	}

	return resp.StatusCode, "", nil
}

func findTitle(n *html.Node) string { // 独立 DFS 函数
	if n.Type == html.ElementNode && n.Data == "title" {
		if n.FirstChild != nil {
			return n.FirstChild.Data
		}
	}
	for c := n.FirstChild; c != nil; c = c.NextSibling {
		if title := findTitle(c); title != "" {
			return title
		}
	}
	return ""
}

func worker(config *Config, urls <-chan string, results chan<- UrlResult, stat *Stats, wg *sync.WaitGroup, stopChan <-chan struct{}) {
	defer wg.Done()
	for url := range urls {
		select {
		case <-stopChan:
			return
		default:
			status_code, title, err := fetchUrl(url, config.Timeout, config.Tls)

			result := UrlResult{
				Url:    url,
				Status: status_code,
				Title:  title,
				Error:  err,
			}
			atomic.AddInt32(&stat.Url, 1)
			switch status_code {
			case 200:
				atomic.AddInt32(&stat.Success, 1)
			case 403:
				atomic.AddInt32(&stat.ForBidden, 1)
			case 404:
				atomic.AddInt32(&stat.NotFound, 1)
			case -1:
				atomic.AddInt32(&stat.Timeout, 1)
			default:
				atomic.AddInt32(&stat.Failed, 1)
			}

			results <- result

		}

	}
}
func WriteResult(stat *Stats, results <-chan UrlResult, config *Config, done chan<- bool) {
	defer func() {
		done <- true
	}()
	if config.OutFile != "" {
		f := excelize.NewFile()
		sheet := f.GetSheetName(0)
		headers := []interface{}{"域名", "状态", "标题"}
		f.SetSheetRow(sheet, "A1", &headers)
		row := 2
		for r := range results {
			data := []interface{}{
				r.Url,
				r.Status,
				r.Title,
			}
			cell, _ := excelize.CoordinatesToCellName(1, row)
			f.SetSheetRow(sheet, cell, &data)
			row++
		}

		if err := f.SaveAs(config.OutFile); err != nil {
			log.Fatal("输出Excel格式文件错误")
		}

	} else {
		for r := range results {
			switch r.Status {
			case 200:
				color.Green("[+]Url: %s   Title: %s", r.Url, r.Title)
			case 403, 404:
				color.Yellow("[*]Url: %s  Title: %s", r.Url, r.Title)
			case -1:
				color.White("[!]Url: %s timeout", r.Url)
			default:
				color.Red("[-]Url: %s connect error", r.Url)
			}
		}

	}
	done <- true

}

func printStats(stat *Stats) {
	fmt.Printf("总url数: %d\n", stat.Url)
	fmt.Printf("状态200数: %d\n", stat.Success)
	fmt.Printf("状态403数: %d\n", stat.ForBidden)
	fmt.Printf("状态404数: %d\n", stat.NotFound)
	fmt.Printf("访问超时数: %d\n", stat.Timeout)
	fmt.Printf("访问失败数: %d\n", stat.Failed)

}
