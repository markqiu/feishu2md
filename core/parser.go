package core

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"strings"

	"github.com/Wsine/feishu2md/utils"
	"github.com/chyroc/lark"
	"github.com/olekukonko/tablewriter"
)

type Parser struct {
	client     *Client
	useHTMLTags bool
	ImgTokens   []string
	blockMap    map[string]*lark.DocxBlock
	ctx         context.Context
	outputDir   string
}

func NewParser(config OutputConfig, client *Client) *Parser {
	return &Parser{
		client:     client,
		useHTMLTags: config.UseHTMLTags,
		ImgTokens:   make([]string, 0),
		blockMap:    make(map[string]*lark.DocxBlock),
		ctx:         context.Background(),
		outputDir:   "",
	}
}

// SetContext sets the context for the parser
func (p *Parser) SetContext(ctx context.Context) {
	p.ctx = ctx
}

// SetOutputDir sets the output directory for the parser
func (p *Parser) SetOutputDir(outputDir string) {
	p.outputDir = outputDir
}

// =============================================================
// Parser utils
// =============================================================

var DocxCodeLang2MdStr = map[lark.DocxCodeLanguage]string{
	lark.DocxCodeLanguagePlainText:    "",
	lark.DocxCodeLanguageABAP:         "abap",
	lark.DocxCodeLanguageAda:          "ada",
	lark.DocxCodeLanguageApache:       "apache",
	lark.DocxCodeLanguageApex:         "apex",
	lark.DocxCodeLanguageAssembly:     "assembly",
	lark.DocxCodeLanguageBash:         "bash",
	lark.DocxCodeLanguageCSharp:       "csharp",
	lark.DocxCodeLanguageCPlusPlus:    "cpp",
	lark.DocxCodeLanguageC:            "c",
	lark.DocxCodeLanguageCOBOL:        "cobol",
	lark.DocxCodeLanguageCSS:          "css",
	lark.DocxCodeLanguageCoffeeScript: "coffeescript",
	lark.DocxCodeLanguageD:            "d",
	lark.DocxCodeLanguageDart:         "dart",
	lark.DocxCodeLanguageDelphi:       "delphi",
	lark.DocxCodeLanguageDjango:       "django",
	lark.DocxCodeLanguageDockerfile:   "dockerfile",
	lark.DocxCodeLanguageErlang:       "erlang",
	lark.DocxCodeLanguageFortran:      "fortran",
	lark.DocxCodeLanguageFoxPro:       "foxpro",
	lark.DocxCodeLanguageGo:           "go",
	lark.DocxCodeLanguageGroovy:       "groovy",
	lark.DocxCodeLanguageHTML:         "html",
	lark.DocxCodeLanguageHTMLBars:     "htmlbars",
	lark.DocxCodeLanguageHTTP:         "http",
	lark.DocxCodeLanguageHaskell:      "haskell",
	lark.DocxCodeLanguageJSON:         "json",
	lark.DocxCodeLanguageJava:         "java",
	lark.DocxCodeLanguageJavaScript:   "javascript",
	lark.DocxCodeLanguageJulia:        "julia",
	lark.DocxCodeLanguageKotlin:       "kotlin",
	lark.DocxCodeLanguageLateX:        "latex",
	lark.DocxCodeLanguageLisp:         "lisp",
	lark.DocxCodeLanguageLogo:         "logo",
	lark.DocxCodeLanguageLua:          "lua",
	lark.DocxCodeLanguageMATLAB:       "matlab",
	lark.DocxCodeLanguageMakefile:     "makefile",
	lark.DocxCodeLanguageMarkdown:     "markdown",
	lark.DocxCodeLanguageNginx:        "nginx",
	lark.DocxCodeLanguageObjective:    "objectivec",
	lark.DocxCodeLanguageOpenEdgeABL:  "openedge-abl",
	lark.DocxCodeLanguagePHP:          "php",
	lark.DocxCodeLanguagePerl:         "perl",
	lark.DocxCodeLanguagePostScript:   "postscript",
	lark.DocxCodeLanguagePower:        "powershell",
	lark.DocxCodeLanguageProlog:       "prolog",
	lark.DocxCodeLanguageProtoBuf:     "protobuf",
	lark.DocxCodeLanguagePython:       "python",
	lark.DocxCodeLanguageR:            "r",
	lark.DocxCodeLanguageRPG:          "rpg",
	lark.DocxCodeLanguageRuby:         "ruby",
	lark.DocxCodeLanguageRust:         "rust",
	lark.DocxCodeLanguageSAS:          "sas",
	lark.DocxCodeLanguageSCSS:         "scss",
	lark.DocxCodeLanguageSQL:          "sql",
	lark.DocxCodeLanguageScala:        "scala",
	lark.DocxCodeLanguageScheme:       "scheme",
	lark.DocxCodeLanguageScratch:      "scratch",
	lark.DocxCodeLanguageShell:        "shell",
	lark.DocxCodeLanguageSwift:        "swift",
	lark.DocxCodeLanguageThrift:       "thrift",
	lark.DocxCodeLanguageTypeScript:   "typescript",
	lark.DocxCodeLanguageVBScript:     "vbscript",
	lark.DocxCodeLanguageVisual:       "vbnet",
	lark.DocxCodeLanguageXML:          "xml",
	lark.DocxCodeLanguageYAML:         "yaml",
}

func renderMarkdownTable(data [][]string) string {
	builder := &strings.Builder{}
	table := tablewriter.NewWriter(builder)
	table.SetCenterSeparator("|")
	table.SetAutoWrapText(false)
	table.SetAutoFormatHeaders(false)
	table.SetAutoMergeCells(false)
	table.SetBorders(tablewriter.Border{Left: true, Top: false, Right: true, Bottom: false})
	table.SetHeader(data[0])
	table.AppendBulk(data[1:])
	table.Render()
	return builder.String()
}

// =============================================================
// Parse the new version of document (docx)
// =============================================================

func (p *Parser) ParseDocxContent(doc *lark.DocxDocument, blocks []*lark.DocxBlock) string {
	for _, block := range blocks {
		p.blockMap[block.BlockID] = block
	}

	entryBlock := p.blockMap[doc.DocumentID]
	return p.ParseDocxBlock(entryBlock, 0)
}

func (p *Parser) ParseDocxBlock(b *lark.DocxBlock, indentLevel int) string {
	buf := new(strings.Builder)
	buf.WriteString(strings.Repeat("\t", indentLevel))

	switch b.BlockType {
	case lark.DocxBlockTypePage:
		buf.WriteString(p.ParseDocxBlockPage(b))
	case lark.DocxBlockTypeText:
		buf.WriteString(p.ParseDocxBlockText(b.Text))
	case lark.DocxBlockTypeCallout:
		buf.WriteString(p.ParseDocxBlockCallout(b))
	case lark.DocxBlockTypeHeading1:
		buf.WriteString(p.ParseDocxBlockHeading(b, 1))
	case lark.DocxBlockTypeHeading2:
		buf.WriteString(p.ParseDocxBlockHeading(b, 2))
	case lark.DocxBlockTypeHeading3:
		buf.WriteString(p.ParseDocxBlockHeading(b, 3))
	case lark.DocxBlockTypeHeading4:
		buf.WriteString(p.ParseDocxBlockHeading(b, 4))
	case lark.DocxBlockTypeHeading5:
		buf.WriteString(p.ParseDocxBlockHeading(b, 5))
	case lark.DocxBlockTypeHeading6:
		buf.WriteString(p.ParseDocxBlockHeading(b, 6))
	case lark.DocxBlockTypeHeading7:
		buf.WriteString(p.ParseDocxBlockHeading(b, 7))
	case lark.DocxBlockTypeHeading8:
		buf.WriteString(p.ParseDocxBlockHeading(b, 8))
	case lark.DocxBlockTypeHeading9:
		buf.WriteString(p.ParseDocxBlockHeading(b, 9))
	case lark.DocxBlockTypeBullet:
		buf.WriteString(p.ParseDocxBlockBullet(b, indentLevel))
	case lark.DocxBlockTypeOrdered:
		buf.WriteString(p.ParseDocxBlockOrdered(b, indentLevel))
	case lark.DocxBlockTypeCode:
		buf.WriteString("```" + DocxCodeLang2MdStr[b.Code.Style.Language] + "\n")
		buf.WriteString(strings.TrimSpace(p.ParseDocxBlockText(b.Code)))
		buf.WriteString("\n```\n")
	case lark.DocxBlockTypeQuote:
		buf.WriteString("> ")
		buf.WriteString(p.ParseDocxBlockText(b.Quote))
	case lark.DocxBlockTypeEquation:
		buf.WriteString("$$\n")
		buf.WriteString(p.ParseDocxBlockText(b.Equation))
		buf.WriteString("\n$$\n")
	case lark.DocxBlockTypeTodo:
		if b.Todo.Style.Done {
			buf.WriteString("- [x] ")
		} else {
			buf.WriteString("- [ ] ")
		}
		buf.WriteString(p.ParseDocxBlockText(b.Todo))
	case lark.DocxBlockTypeDivider:
		buf.WriteString("---\n")
	case lark.DocxBlockTypeImage:
		buf.WriteString(p.ParseDocxBlockImage(b.Image))
	case lark.DocxBlockTypeFile:
		buf.WriteString(p.ParseDocxBlockFile(b.File))
	case lark.DocxBlockTypeBitable:
		buf.WriteString(p.ParseDocxBlockBitable(b.Bitable))
	case lark.DocxBlockTypeDiagram:
		buf.WriteString(p.ParseDocxBlockDiagram(b.Diagram))
	case lark.DocxBlockTypeIframe:
		buf.WriteString(p.ParseDocxBlockIframe(b.Iframe))
	case lark.DocxBlockTypeTableCell:
		buf.WriteString(p.ParseDocxBlockTableCell(b))
	case lark.DocxBlockTypeTable:
		buf.WriteString(p.ParseDocxBlockTable(b.Table))
	case lark.DocxBlockTypeSheet:
		buf.WriteString(p.ParseDocxBlockSheet(b.Sheet))
	case lark.DocxBlockTypeQuoteContainer:
		buf.WriteString(p.ParseDocxBlockQuoteContainer(b))
	case lark.DocxBlockTypeGrid:
		buf.WriteString(p.ParseDocxBlockGrid(b, indentLevel))
	default:
		// 对于不支持的 block type，仍然处理其 children
		for _, childId := range b.Children {
			childBlock := p.blockMap[childId]
			buf.WriteString(p.ParseDocxBlock(childBlock, indentLevel))
		}
	}
	return buf.String()
}

func (p *Parser) ParseDocxBlockPage(b *lark.DocxBlock) string {
	buf := new(strings.Builder)

	buf.WriteString("# ")
	buf.WriteString(p.ParseDocxBlockText(b.Page))
	buf.WriteString("\n")

	for _, childId := range b.Children {
		childBlock := p.blockMap[childId]
		buf.WriteString(p.ParseDocxBlock(childBlock, 0))
		buf.WriteString("\n")
	}

	return buf.String()
}

func (p *Parser) ParseDocxBlockText(b *lark.DocxBlockText) string {
	buf := new(strings.Builder)
	numElem := len(b.Elements)
	for _, e := range b.Elements {
		inline := numElem > 1
		buf.WriteString(p.ParseDocxTextElement(e, inline))
	}
	buf.WriteString("\n")
	return buf.String()
}

func (p *Parser) ParseDocxBlockCallout(b *lark.DocxBlock) string {
	buf := new(strings.Builder)

	buf.WriteString(">[!TIP] \n")

	for _, childId := range b.Children {
		childBlock := p.blockMap[childId]
		buf.WriteString(p.ParseDocxBlock(childBlock, 0))
	}

	return buf.String()
}
func (p *Parser) ParseDocxTextElement(e *lark.DocxTextElement, inline bool) string {
	buf := new(strings.Builder)
	if e.TextRun != nil {
		buf.WriteString(p.ParseDocxTextElementTextRun(e.TextRun))
	}
	if e.MentionUser != nil {
		buf.WriteString(e.MentionUser.UserID)
	}
	if e.MentionDoc != nil {
		buf.WriteString(
			fmt.Sprintf("[%s](%s)", e.MentionDoc.Title, utils.UnescapeURL(e.MentionDoc.URL)))
	}
	if e.Equation != nil {
		symbol := "$$"
		if inline {
			symbol = "$"
		}
		buf.WriteString(symbol + strings.TrimSuffix(e.Equation.Content, "\n") + symbol)
	}
	return buf.String()
}

func (p *Parser) ParseDocxTextElementTextRun(tr *lark.DocxTextElementTextRun) string {
	buf := new(strings.Builder)
	postWrite := ""
	if style := tr.TextElementStyle; style != nil {
		if style.Bold {
			if p.useHTMLTags {
				buf.WriteString("<strong>")
				postWrite = "</strong>"
			} else {
				buf.WriteString("**")
				postWrite = "**"
			}
		} else if style.Italic {
			if p.useHTMLTags {
				buf.WriteString("<em>")
				postWrite = "</em>"
			} else {
				buf.WriteString("_")
				postWrite = "_"
			}
		} else if style.Strikethrough {
			if p.useHTMLTags {
				buf.WriteString("<del>")
				postWrite = "</del>"
			} else {
				buf.WriteString("~~")
				postWrite = "~~"
			}
		} else if style.Underline {
			buf.WriteString("<u>")
			postWrite = "</u>"
		} else if style.InlineCode {
			buf.WriteString("`")
			postWrite = "`"
		} else if link := style.Link; link != nil {
			buf.WriteString("[")
			postWrite = fmt.Sprintf("](%s)", utils.UnescapeURL(link.URL))
		}
	}
	buf.WriteString(tr.Content)
	buf.WriteString(postWrite)
	return buf.String()
}

func (p *Parser) ParseDocxBlockHeading(b *lark.DocxBlock, headingLevel int) string {
	buf := new(strings.Builder)

	buf.WriteString(strings.Repeat("#", headingLevel))
	buf.WriteString(" ")

	headingText := reflect.ValueOf(b).Elem().FieldByName(fmt.Sprintf("Heading%d", headingLevel))
	buf.WriteString(p.ParseDocxBlockText(headingText.Interface().(*lark.DocxBlockText)))

	for _, childId := range b.Children {
		childBlock := p.blockMap[childId]
		buf.WriteString(p.ParseDocxBlock(childBlock, 0))
	}

	return buf.String()
}

func (p *Parser) ParseDocxBlockImage(img *lark.DocxBlockImage) string {
	buf := new(strings.Builder)
	buf.WriteString(fmt.Sprintf("![](%s)", img.Token))
	buf.WriteString("\n")
	p.ImgTokens = append(p.ImgTokens, img.Token)
	return buf.String()
}

func (p *Parser) ParseDocxBlockFile(file *lark.DocxBlockFile) string {
	buf := new(strings.Builder)
	
	// Get file extension to determine file type
	var fileType string
	var fileName string
	if file.Name != "" {
		fileName = file.Name
	} else {
		fileName = file.Token
	}
	
	// Determine file type based on name or token
	if strings.Contains(strings.ToLower(fileName), ".mp4") || 
	   strings.Contains(strings.ToLower(fileName), ".mov") ||
	   strings.Contains(strings.ToLower(fileName), ".avi") ||
	   strings.Contains(strings.ToLower(fileName), ".mkv") {
		fileType = "视频"
	} else if strings.Contains(strings.ToLower(fileName), ".pdf") {
		fileType = "PDF"
	} else if strings.Contains(strings.ToLower(fileName), ".doc") ||
	   strings.Contains(strings.ToLower(fileName), ".docx") {
		fileType = "Word文档"
	} else if strings.Contains(strings.ToLower(fileName), ".xls") ||
	   strings.Contains(strings.ToLower(fileName), ".xlsx") {
		fileType = "Excel表格"
	} else {
		fileType = "文件"
	}
	
	buf.WriteString(fmt.Sprintf("\n**附件**: %s (%s)\n\n", fileName, fileType))
	
	// Try to download the file if context and outputDir are set
	// For file blocks inside documents, we should use DownloadDriveMedia
	if p.ctx != nil && p.outputDir != "" && p.client != nil {
		// Use DownloadDriveMedia for file blocks inside documents
		resp, _, err := p.client.larkClient.Drive.DownloadDriveMedia(p.ctx, &lark.DownloadDriveMediaReq{
			FileToken: file.Token,
		})
		
		if err == nil && resp != nil {
			// File downloaded successfully
			downloadedFilename := resp.Filename
			if downloadedFilename == "" {
				downloadedFilename = file.Token
			}
			
			filePath := filepath.Join(p.outputDir, downloadedFilename)
			err := os.MkdirAll(filepath.Dir(filePath), 0o755)
			if err == nil {
				file, err := os.OpenFile(filePath, os.O_CREATE|os.O_WRONLY, 0o666)
				if err == nil {
					written, err := file.ReadFrom(resp.File)
					if err == nil {
						buf.WriteString(fmt.Sprintf("**下载成功**: 文件已保存到 `%s` (大小: %d bytes)\n\n", filePath, written))
						return buf.String()
					}
				}
			}
		}
		// Download failed, fall through to placeholder
	}
	
	buf.WriteString(fmt.Sprintf("**文件Token**: `%s`\n\n", file.Token))
	buf.WriteString(fmt.Sprintf("**提示**: 这是一个%s附件，请访问飞书查看原始文件。\n\n", fileType))
	
	return buf.String()
}

func (p *Parser) ParseDocxWhatever(body *lark.DocBody) string {
	buf := new(strings.Builder)

	return buf.String()
}

func (p *Parser) ParseDocxBlockBullet(b *lark.DocxBlock, indentLevel int) string {
	buf := new(strings.Builder)

	buf.WriteString("- ")
	buf.WriteString(p.ParseDocxBlockText(b.Bullet))

	for _, childId := range b.Children {
		childBlock := p.blockMap[childId]
		buf.WriteString(p.ParseDocxBlock(childBlock, indentLevel+1))
	}

	return buf.String()
}

func (p *Parser) ParseDocxBlockOrdered(b *lark.DocxBlock, indentLevel int) string {
	buf := new(strings.Builder)

	// calculate order and indent level
	parent := p.blockMap[b.ParentID]
	order := 1
	for idx, child := range parent.Children {
		if child == b.BlockID {
			for i := idx - 1; i >= 0; i-- {
				if p.blockMap[parent.Children[i]].BlockType == lark.DocxBlockTypeOrdered {
					order += 1
				} else {
					break
				}
			}
			break
		}
	}

	buf.WriteString(fmt.Sprintf("%d. ", order))
	buf.WriteString(p.ParseDocxBlockText(b.Ordered))

	for _, childId := range b.Children {
		childBlock := p.blockMap[childId]
		buf.WriteString(p.ParseDocxBlock(childBlock, indentLevel+1))
	}

	return buf.String()
}

func (p *Parser) ParseDocxBlockTableCell(b *lark.DocxBlock) string {
	buf := new(strings.Builder)

	for _, child := range b.Children {
		block := p.blockMap[child]
		content := p.ParseDocxBlock(block, 0)
		buf.WriteString(content + "<br/>")
	}

	return buf.String()
}

func (p *Parser) ParseDocxBlockTable(t *lark.DocxBlockTable) string {
	var rows [][]string
	mergeInfoMap := map[int64]map[int64]*lark.DocxBlockTablePropertyMergeInfo{}

	// 构建单元格合并信息的映射
	if t.Property.MergeInfo != nil {
		for i, merge := range t.Property.MergeInfo {
			rowIndex := int64(i) / t.Property.ColumnSize
			colIndex := int64(i) % t.Property.ColumnSize
			if _, exists := mergeInfoMap[int64(rowIndex)]; !exists {
				mergeInfoMap[int64(rowIndex)] = map[int64]*lark.DocxBlockTablePropertyMergeInfo{}
			}
			mergeInfoMap[rowIndex][colIndex] = merge
		}
	}

	// 构建表格内容

	for i, blockId := range t.Cells {
		block := p.blockMap[blockId]
		cellContent := p.ParseDocxBlock(block, 0)
		cellContent = strings.ReplaceAll(cellContent, "\n", "")
		rowIndex := int64(i) / t.Property.ColumnSize
		colIndex := int64(i) % t.Property.ColumnSize

		// 初始化行
		for len(rows) <= int(rowIndex) {
			rows = append(rows, []string{})
		}
		for len(rows[rowIndex]) <= int(colIndex) {
			rows[rowIndex] = append(rows[rowIndex], "")
		}
		// 设置单元格内容
		rows[rowIndex][colIndex] = cellContent
	}

	// 渲染为 HTML 表格
	buf := new(strings.Builder)
	buf.WriteString("<table>\n")

	// 跟踪已经处理过的合并单元格
	processedCells := map[string]bool{}

	// 构建 HTML 表格内容
	for rowIndex, row := range rows {
		buf.WriteString("<tr>\n")
		for colIndex, cellContent := range row {
			cellKey := fmt.Sprintf("%d-%d", rowIndex, colIndex)

			// 跳过已处理的单元格
			if processedCells[cellKey] {
				continue
			}

			mergeInfo := mergeInfoMap[int64(rowIndex)][int64(colIndex)]
			if mergeInfo != nil {

				// 合并单元格，只有当 RowSpan > 1 或 ColSpan > 1 时才添加对应属性
				attributes := ""
				if mergeInfo.RowSpan > 1 {
					attributes += fmt.Sprintf(` rowspan="%d"`, mergeInfo.RowSpan)
				}
				if mergeInfo.ColSpan > 1 {
					attributes += fmt.Sprintf(` colspan="%d"`, mergeInfo.ColSpan)
				}
				buf.WriteString(fmt.Sprintf(
					`<td%s>%s</td>`,
					attributes, cellContent,
				))
				// 标记合并范围内的所有单元格为已处理
				for r := rowIndex; r < rowIndex+int(mergeInfo.RowSpan); r++ {
					for c := colIndex; c < colIndex+int(mergeInfo.ColSpan); c++ {
						processedCells[fmt.Sprintf("%d-%d", r, c)] = true
					}
				}
			} else {
				// 普通单元格
				buf.WriteString(fmt.Sprintf("<td>%s</td>", cellContent))
			}
		}
		buf.WriteString("</tr>\n")
	}
	buf.WriteString("</table>\n")

	return buf.String()
}

func (p *Parser) ParseDocxBlockQuoteContainer(b *lark.DocxBlock) string {
	buf := new(strings.Builder)

	for i, child := range b.Children {
		block := p.blockMap[child]
		buf.WriteString("> ")
		content := p.ParseDocxBlock(block, 0)
		// 移除内容末尾的换行符
		content = strings.TrimRight(content, "\n")
		buf.WriteString(content)
		// 在行尾添加两个空格来实现换行（markdown 语法）
		buf.WriteString("  ")
		// 如果不是最后一个子块，则添加换行符
		if i < len(b.Children)-1 {
			buf.WriteString("\n")
		}
	}

	return buf.String()
}

func (p *Parser) ParseDocxBlockGrid(b *lark.DocxBlock, indentLevel int) string {
	buf := new(strings.Builder)

	for _, child := range b.Children {
		columnBlock := p.blockMap[child]
		for _, child := range columnBlock.Children {
			block := p.blockMap[child]
			buf.WriteString(p.ParseDocxBlock(block, indentLevel))
		}
	}

	return buf.String()
}

func (p *Parser) ParseDocxBlockSheet(s *lark.DocxBlockSheet) string {
	// 电子表格块（Sheet）是嵌入到飞书文档中的外部电子表格
	buf := new(strings.Builder)

	// 如果没有 client 或 token，则返回占位符
	if p.client == nil || s.Token == "" {
		buf.WriteString("\n\n")
		buf.WriteString("> **📊 嵌入的电子表格**\n")
		buf.WriteString(">\n")
		if s.Token != "" {
			buf.WriteString(fmt.Sprintf("> Token: `%s`\n", s.Token))
		}
		buf.WriteString(">\n")
		buf.WriteString("> *注：无法获取电子表格内容（缺少 client 或 token）*\n")
		buf.WriteString("\n\n")
		return buf.String()
	}

	// 尝试获取电子表格的实际内容
	ctx := context.Background()
	values, err := p.client.GetSheetContent(ctx, s.Token)
	if err != nil {
		// 如果获取失败，返回占位符
		buf.WriteString("\n\n")
		buf.WriteString("> **📊 嵌入的电子表格**\n")
		buf.WriteString(">\n")
		if s.Token != "" {
			buf.WriteString(fmt.Sprintf("> Token: `%s`\n", s.Token))
		}
		buf.WriteString(">\n")
		// 检查是否是 token 格式问题
		if strings.Contains(err.Error(), "invalid spreadsheet token format") {
			buf.WriteString("> *注：此电子表格使用了不支持的嵌入方式，无法获取内容*\n")
		} else if strings.Contains(err.Error(), "91402") || strings.Contains(err.Error(), "NOTEXIST") {
			buf.WriteString("> *注：无法访问电子表格（可能没有权限或电子表格不存在）*\n")
		} else {
			buf.WriteString(fmt.Sprintf("> *获取电子表格内容失败: %v*\n", err))
		}
		buf.WriteString("\n\n")
		return buf.String()
	}

	// 将电子表格数据转换为 markdown 表格
	if len(values) == 0 {
		buf.WriteString("\n\n")
		buf.WriteString("> **📊 嵌入的电子表格**\n")
		buf.WriteString(">\n")
		if s.Token != "" {
			buf.WriteString(fmt.Sprintf("> Token: `%s`\n", s.Token))
		}
		buf.WriteString(">\n")
		buf.WriteString("> *电子表格为空*\n")
		buf.WriteString("\n\n")
		return buf.String()
	}

	// 生成 markdown 表格
	buf.WriteString("\n\n")
	// 表头
	buf.WriteString("|")
	for _, cell := range values[0] {
		buf.WriteString(" " + cell + " |")
	}
	buf.WriteString("\n")
	// 分隔线
	buf.WriteString("|")
	for range values[0] {
		buf.WriteString(" --- |")
	}
	buf.WriteString("\n")
	// 数据行
	for i := 1; i < len(values); i++ {
		buf.WriteString("|")
		for _, cell := range values[i] {
			buf.WriteString(" " + cell + " |")
		}
		buf.WriteString("\n")
	}
	buf.WriteString("\n")

	return buf.String()
}

// ParseDocxBlockBitable 解析多维表格块
func (p *Parser) ParseDocxBlockBitable(bitable *lark.DocxBlockBitable) string {
	buf := new(strings.Builder)

	// 如果没有 client 或 token，则返回占位符
	if p.client == nil || bitable.Token == "" {
		buf.WriteString("\n\n")
		buf.WriteString("> **📊 多维表格**\n")
		buf.WriteString(">\n")
		if bitable.Token != "" {
			buf.WriteString(fmt.Sprintf("> Token: `%s`\n", bitable.Token))
		}
		buf.WriteString(">\n")
		buf.WriteString("> *注：无法获取多维表格内容（缺少 client 或 token）*\n")
		buf.WriteString("\n\n")
		return buf.String()
	}

	// 尝试获取多维表格的实际内容
	ctx := context.Background()
	values, err := p.client.GetBitableContent(ctx, bitable.Token)
	if err != nil {
		// 如果获取失败，返回占位符
		buf.WriteString("\n\n")
		buf.WriteString("> **📊 多维表格**\n")
		buf.WriteString(">\n")
		if bitable.Token != "" {
			buf.WriteString(fmt.Sprintf("> Token: `%s`\n", bitable.Token))
		}
		buf.WriteString(">\n")
		buf.WriteString(fmt.Sprintf("> *获取多维表格内容失败: %v*\n", err))
		buf.WriteString("\n\n")
		return buf.String()
	}

	// 将多维表格数据转换为 markdown 表格
	if len(values) == 0 {
		buf.WriteString("\n\n")
		buf.WriteString("> **📊 多维表格**\n")
		buf.WriteString(">\n")
		if bitable.Token != "" {
			buf.WriteString(fmt.Sprintf("> Token: `%s`\n", bitable.Token))
		}
		buf.WriteString(">\n")
		buf.WriteString("> *多维表格为空*\n")
		buf.WriteString("\n\n")
		return buf.String()
	}

	// 生成 markdown 表格
	buf.WriteString("\n\n")
	// 表头
	buf.WriteString("|")
	for _, cell := range values[0] {
		buf.WriteString(" " + cell + " |")
	}
	buf.WriteString("\n")
	// 分隔线
	buf.WriteString("|")
	for range values[0] {
		buf.WriteString(" --- |")
	}
	buf.WriteString("\n")
	// 数据行
	for i := 1; i < len(values); i++ {
		buf.WriteString("|")
		for _, cell := range values[i] {
			buf.WriteString(" " + cell + " |")
		}
		buf.WriteString("\n")
	}
	buf.WriteString("\n")

	return buf.String()
}

// ParseDocxBlockDiagram 解析流程图/UML块
func (p *Parser) ParseDocxBlockDiagram(diagram *lark.DocxBlockDiagram) string {
	buf := new(strings.Builder)

	diagramType := "流程图"
	if diagram.DiagramType == 2 {
		diagramType = "UML图"
	}

	buf.WriteString("\n\n")
	buf.WriteString(fmt.Sprintf("**📈 %s**\n\n", diagramType))
	buf.WriteString("> *注：流程图/UML图无法直接转换为 Markdown，建议导出为图片或使用 Mermaid 语法*\n")
	buf.WriteString("\n\n")

	return buf.String()
}

// ParseDocxBlockIframe 解析内嵌块
func (p *Parser) ParseDocxBlockIframe(iframe *lark.DocxBlockIframe) string {
	buf := new(strings.Builder)

	buf.WriteString("\n\n")
	buf.WriteString("**🔗 嵌入内容**\n\n")

	if iframe.Component != nil {
		// 获取 iframe 类型名称
		typeNames := map[int]string{
			1:  "哔哩哔哩",
			2:  "西瓜视频",
			3:  "优酷",
			4:  "Airtable",
			5:  "百度地图",
			6:  "高德地图",
			7:  "TikTok",
			8:  "Figma",
			9:  "墨刀",
			10: "Canva",
			11: "CodePen",
			12: "飞书问卷",
			13: "金数据",
			14: "谷歌地图",
			15: "YouTube",
			99: "其他",
		}

		typeName := "未知类型"
		if name, ok := typeNames[int(iframe.Component.IframeType)]; ok {
			typeName = name
		}

		buf.WriteString(fmt.Sprintf("> 类型: %s\n", typeName))

		// 显示 URL（如果有的话）
		if iframe.Component.URL != "" {
			buf.WriteString(">\n")
			buf.WriteString(fmt.Sprintf("> 链接: %s\n", iframe.Component.URL))
		}
	}

	buf.WriteString(">\n")
	buf.WriteString("> *注：嵌入内容无法直接在 Markdown 中显示，请访问飞书查看原始内容*\n")
	buf.WriteString("\n\n")

	return buf.String()
}
