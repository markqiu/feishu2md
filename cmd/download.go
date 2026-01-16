package main

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"sync"

	"github.com/88250/lute"
	"github.com/Wsine/feishu2md/core"
	"github.com/Wsine/feishu2md/utils"
	"github.com/chyroc/lark"
	"github.com/pkg/errors"
)

type DownloadOpts struct {
	outputDir string
	dump      bool
	batch     bool
	wiki      bool
}

var dlOpts = DownloadOpts{}
var dlConfig core.Config

func downloadDocument(ctx context.Context, client *core.Client, url string, opts *DownloadOpts) error {
	// Validate the url to download
	docType, docToken, err := utils.ValidateDocumentURL(url)
	if err != nil {
		return err
	}
	fmt.Println("Captured document token:", docToken)

	// for a wiki page, we need to renew docType and docToken first
	var nodeTitle string
	if docType == "wiki" {
		node, err := client.GetWikiNodeInfo(ctx, docToken)
		if err != nil {
			err = fmt.Errorf("GetWikiNodeInfo err: %v for %v", err, url)
		}
		utils.CheckErr(err)
		docType = node.ObjType
		docToken = node.ObjToken
		nodeTitle = node.Title
	}
	if docType == "docs" {
		return errors.Errorf(
			`Feishu Docs is no longer supported. ` +
				`Please refer to the Readme/Release for v1_support.`)
	}

	// Handle non-docx file types (mindnote, file, sheet, bitable)
	if docType != "docx" {
		return downloadFile(ctx, client, docToken, nodeTitle, opts.outputDir, docType)
	}

	// Process the download
	docx, blocks, err := client.GetDocxContent(ctx, docToken)
	utils.CheckErr(err)

	parser := core.NewParser(dlConfig.Output, client)
	parser.SetContext(ctx)
	parser.SetOutputDir(filepath.Join(opts.outputDir, dlConfig.Output.ImageDir))

	title := docx.Title
	markdown := parser.ParseDocxContent(docx, blocks)

	if !dlConfig.Output.SkipImgDownload {
		for _, imgToken := range parser.ImgTokens {
			localLink, err := client.DownloadImage(
				ctx, imgToken, filepath.Join(opts.outputDir, dlConfig.Output.ImageDir),
			)
			if err != nil {
				return err
			}
			markdown = strings.Replace(markdown, imgToken, localLink, 1)
		}
	}

	// Format the markdown document
	engine := lute.New(func(l *lute.Lute) {
		l.RenderOptions.AutoSpace = true
	})
	result := engine.FormatStr("md", markdown)

	// Handle the output directory and name
	if _, err := os.Stat(opts.outputDir); os.IsNotExist(err) {
		if err := os.MkdirAll(opts.outputDir, 0o755); err != nil {
			return err
		}
	}

	if dlOpts.dump {
		jsonName := fmt.Sprintf("%s.json", docToken)
		outputPath := filepath.Join(opts.outputDir, jsonName)
		data := struct {
			Document *lark.DocxDocument `json:"document"`
			Blocks   []*lark.DocxBlock  `json:"blocks"`
		}{
			Document: docx,
			Blocks:   blocks,
		}
		pdata := utils.PrettyPrint(data)

		if err = os.WriteFile(outputPath, []byte(pdata), 0o644); err != nil {
			return err
		}
		fmt.Printf("Dumped json response to %s\n", outputPath)
	}

	// Write to markdown file
	mdName := fmt.Sprintf("%s.md", docToken)
	if dlConfig.Output.TitleAsFilename {
		mdName = fmt.Sprintf("%s.md", utils.SanitizeFileName(title))
	}
	outputPath := filepath.Join(opts.outputDir, mdName)
	if err = os.WriteFile(outputPath, []byte(result), 0o644); err != nil {
		return err
	}
	fmt.Printf("Downloaded markdown file to %s\n", outputPath)

	return nil
}

func downloadDocuments(ctx context.Context, client *core.Client, url string) error {
	// Validate the url to download
	folderToken, err := utils.ValidateFolderURL(url)
	if err != nil {
		return err
	}
	fmt.Println("Captured folder token:", folderToken)

	// Error channel and wait group
	errChan := make(chan error)
	wg := sync.WaitGroup{}

	// Recursively go through the folder and download the documents
	var processFolder func(ctx context.Context, folderPath, folderToken string) error
	processFolder = func(ctx context.Context, folderPath, folderToken string) error {
		files, err := client.GetDriveFolderFileList(ctx, nil, &folderToken)
		if err != nil {
			return err
		}
		opts := DownloadOpts{outputDir: folderPath, dump: dlOpts.dump, batch: false}
		for _, file := range files {
			if file.Type == "folder" {
				_folderPath := filepath.Join(folderPath, file.Name)
				if err := processFolder(ctx, _folderPath, file.Token); err != nil {
					return err
				}
			} else if file.Type == "docx" {
				// concurrently download the document
				wg.Add(1)
				go func(_url string) {
					if err := downloadDocument(ctx, client, _url, &opts); err != nil {
						errChan <- err
					}
					wg.Done()
				}(file.URL)
			}
		}
		return nil
	}
	if err := processFolder(ctx, dlOpts.outputDir, folderToken); err != nil {
		return err
	}

	// Wait for all the downloads to finish
	go func() {
		wg.Wait()
		close(errChan)
	}()
	for err := range errChan {
		return err
	}
	return nil
}

func downloadWiki(ctx context.Context, client *core.Client, url string) error {
	prefixURL, wikiToken, err := utils.ValidateWikiURL(url)
	if err != nil {
		return err
	}

	var spaceID string
	// Check if the token is a space_id (from /wiki/settings/ URL) or a node_token (from /wiki/ URL)
	// Try to get wiki space info first - if it works, it's a space_id
	_, err = client.GetWikiName(ctx, wikiToken)
	if err == nil {
		// It's a valid space_id
		spaceID = wikiToken
	} else {
		// It's likely a node_token, get node info to extract space_id
		node, err := client.GetWikiNodeInfo(ctx, wikiToken)
		if err != nil {
			return fmt.Errorf("failed to get wiki node info: %v", err)
		}
		if node.SpaceID == "" {
			return fmt.Errorf("node does not have a space_id")
		}
		spaceID = node.SpaceID
	}

	folderPath, err := client.GetWikiName(ctx, spaceID)
	if err != nil {
		return err
	}
	if folderPath == "" {
		return fmt.Errorf("failed to GetWikiName")
	}
	// Combine with output directory
	folderPath = filepath.Join(dlOpts.outputDir, folderPath)

	errChan := make(chan error)

	var maxConcurrency = 10 // Set the maximum concurrency level
	wg := sync.WaitGroup{}
	semaphore := make(chan struct{}, maxConcurrency) // Create a semaphore with the maximum concurrency level

	var downloadWikiNode func(ctx context.Context,
		client *core.Client,
		spaceID string,
		parentPath string,
		parentNodeToken *string) error

	downloadWikiNode = func(ctx context.Context,
		client *core.Client,
		spaceID string,
		folderPath string,
		parentNodeToken *string) error {
		nodes, err := client.GetWikiNodeList(ctx, spaceID, parentNodeToken)
		if err != nil {
			return err
		}
		for _, n := range nodes {
			// 先处理节点本身的文档内容（如果有的话）
			// Handle different object types
			if n.ObjType == "docx" {
				opts := DownloadOpts{outputDir: folderPath, dump: dlOpts.dump, batch: false}
				wg.Add(1)
				semaphore <- struct{}{}
				go func(_url string) {
					if err := downloadDocument(ctx, client, _url, &opts); err != nil {
						errChan <- err
					}
					wg.Done()
					<-semaphore
				}(prefixURL + "/wiki/" + n.NodeToken)
			} else if n.ObjType == "mindnote" || n.ObjType == "file" || n.ObjType == "sheet" || n.ObjType == "bitable" {
				// Download other file types (mindnote, video, sheet, bitable, etc.)
				// Capture variables for goroutine
				objToken := n.ObjToken
				title := n.Title
				objType := n.ObjType
				wg.Add(1)
				semaphore <- struct{}{}
				go func() {
					if err := downloadFile(ctx, client, objToken, title, folderPath, objType); err != nil {
						errChan <- err
					}
					wg.Done()
					<-semaphore
				}()
			}
			
			// 然后递归处理子节点
			if n.HasChild {
				_folderPath := filepath.Join(folderPath, n.Title)
				if err := downloadWikiNode(ctx, client,
					spaceID, _folderPath, &n.NodeToken); err != nil {
					return err
				}
			}
		}
		return nil
	}

	if err = downloadWikiNode(ctx, client, spaceID, folderPath, nil); err != nil {
		return err
	}

	// Wait for all the downloads to finish
	go func() {
		wg.Wait()
		close(errChan)
	}()
	for err := range errChan {
		return err
	}
	return nil
}

func handleDownloadCommand(url string) error {
	// Load config
	configPath, err := core.GetConfigFilePath()
	if err != nil {
		return err
	}
	config, err := core.ReadConfigFromFile(configPath)
	if err != nil {
		return err
	}
	dlConfig = *config

	// Instantiate the client
	client := core.NewClient(
		dlConfig.Feishu.AppId, dlConfig.Feishu.AppSecret,
	)
	ctx := context.Background()

	if dlOpts.batch {
		return downloadDocuments(ctx, client, url)
	}

	if dlOpts.wiki {
		return downloadWiki(ctx, client, url)
	}

	return downloadDocument(ctx, client, url, &dlOpts)
}

func downloadFile(ctx context.Context, client *core.Client, nodeToken, title, outputDir, objType string) error {
	// Download the file using the objToken
	filePath, err := client.DownloadFile(ctx, nodeToken, outputDir, objType, title)
	if err != nil {
		return fmt.Errorf("failed to download file %s: %v", title, err)
	}
	fmt.Printf("Downloaded file to %s\n", filePath)
	return nil
}
