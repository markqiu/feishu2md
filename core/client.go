package core

import (
	"bytes"
	"context"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/chyroc/lark"
	"github.com/chyroc/lark_rate_limiter"
)

type Client struct {
	larkClient *lark.Lark
}

func NewClient(appID, appSecret string) *Client {
	return &Client{
		larkClient: lark.New(
			lark.WithAppCredential(appID, appSecret),
			lark.WithTimeout(60*time.Second),
			lark.WithApiMiddleware(lark_rate_limiter.Wait(4, 4)),
		),
	}
}

func (c *Client) DownloadImage(ctx context.Context, imgToken, outDir string) (string, error) {
	resp, _, err := c.larkClient.Drive.DownloadDriveMedia(ctx, &lark.DownloadDriveMediaReq{
		FileToken: imgToken,
	})
	if err != nil {
		return imgToken, err
	}
	fileext := filepath.Ext(resp.Filename)
	filename := fmt.Sprintf("%s/%s%s", outDir, imgToken, fileext)
	err = os.MkdirAll(filepath.Dir(filename), 0o755)
	if err != nil {
		return imgToken, err
	}
	file, err := os.OpenFile(filename, os.O_CREATE|os.O_WRONLY, 0o666)
	if err != nil {
		return imgToken, err
	}
	defer file.Close()
	_, err = io.Copy(file, resp.File)
	if err != nil {
		return imgToken, err
	}
	return filename, nil
}

func (c *Client) DownloadImageRaw(ctx context.Context, imgToken, imgDir string) (string, []byte, error) {
	resp, _, err := c.larkClient.Drive.DownloadDriveMedia(ctx, &lark.DownloadDriveMediaReq{
		FileToken: imgToken,
	})
	if err != nil {
		return imgToken, nil, err
	}
	fileext := filepath.Ext(resp.Filename)
	filename := fmt.Sprintf("%s/%s%s", imgDir, imgToken, fileext)
	buf := new(bytes.Buffer)
	buf.ReadFrom(resp.File)
	return filename, buf.Bytes(), nil
}

func (c *Client) GetDocxContent(ctx context.Context, docToken string) (*lark.DocxDocument, []*lark.DocxBlock, error) {
	resp, _, err := c.larkClient.Drive.GetDocxDocument(ctx, &lark.GetDocxDocumentReq{
		DocumentID: docToken,
	})
	if err != nil {
		return nil, nil, err
	}
	docx := &lark.DocxDocument{
		DocumentID: resp.Document.DocumentID,
		RevisionID: resp.Document.RevisionID,
		Title:      resp.Document.Title,
	}
	var blocks []*lark.DocxBlock
	var pageToken *string
	for {
		resp2, _, err := c.larkClient.Drive.GetDocxBlockListOfDocument(ctx, &lark.GetDocxBlockListOfDocumentReq{
			DocumentID: docx.DocumentID,
			PageToken:  pageToken,
		})
		if err != nil {
			return docx, nil, err
		}
		blocks = append(blocks, resp2.Items...)
		pageToken = &resp2.PageToken
		if !resp2.HasMore {
			break
		}
	}
	return docx, blocks, nil
}

func (c *Client) GetWikiNodeInfo(ctx context.Context, token string) (*lark.GetWikiNodeRespNode, error) {
	resp, _, err := c.larkClient.Drive.GetWikiNode(ctx, &lark.GetWikiNodeReq{
		Token: token,
	})
	if err != nil {
		return nil, err
	}
	return resp.Node, nil
}

func (c *Client) GetDriveFolderFileList(ctx context.Context, pageToken *string, folderToken *string) ([]*lark.GetDriveFileListRespFile, error) {
	resp, _, err := c.larkClient.Drive.GetDriveFileList(ctx, &lark.GetDriveFileListReq{
		PageSize:    nil,
		PageToken:   pageToken,
		FolderToken: folderToken,
	})
	if err != nil {
		return nil, err
	}
	files := resp.Files
	for resp.HasMore {
		resp, _, err = c.larkClient.Drive.GetDriveFileList(ctx, &lark.GetDriveFileListReq{
			PageSize:    nil,
			PageToken:   &resp.NextPageToken,
			FolderToken: folderToken,
		})
		if err != nil {
			return nil, err
		}
		files = append(files, resp.Files...)
	}
	return files, nil
}

func (c *Client) GetWikiName(ctx context.Context, spaceID string) (string, error) {
	resp, _, err := c.larkClient.Drive.GetWikiSpace(ctx, &lark.GetWikiSpaceReq{
		SpaceID: spaceID,
	})

	if err != nil {
		return "", err
	}

	return resp.Space.Name, nil
}

func (c *Client) GetWikiNodeList(ctx context.Context, spaceID string, parentNodeToken *string) ([]*lark.GetWikiNodeListRespItem, error) {
	resp, _, err := c.larkClient.Drive.GetWikiNodeList(ctx, &lark.GetWikiNodeListReq{
		SpaceID:         spaceID,
		PageSize:        nil,
		PageToken:       nil,
		ParentNodeToken: parentNodeToken,
	})

	if err != nil {
		return nil, err
	}

	nodes := resp.Items
	previousPageToken := ""

	for resp.HasMore && previousPageToken != resp.PageToken {
		previousPageToken = resp.PageToken
		resp, _, err := c.larkClient.Drive.GetWikiNodeList(ctx, &lark.GetWikiNodeListReq{
			SpaceID:         spaceID,
			PageSize:        nil,
			PageToken:       &resp.PageToken,
			ParentNodeToken: parentNodeToken,
		})

		if err != nil {
			return nil, err
		}

		nodes = append(nodes, resp.Items...)
	}

	return nodes, nil
}
// GetSheetContent 获取电子表格的内容
func (c *Client) GetSheetContent(ctx context.Context, sheetToken string) ([][]string, error) {
	// sheetToken 的格式是：spreadsheet_token + "_" + sheet_id
	// 例如：B3hasMxsshByaEtZxAwcVfWxnSe_Ml1QzO
	// 需要解析出 spreadsheet_token 和 sheet_id
	
	// 查找最后一个下划线，分隔 spreadsheet_token 和 sheet_id
	lastUnderscore := strings.LastIndex(sheetToken, "_")
	if lastUnderscore == -1 {
		return nil, fmt.Errorf("invalid sheet token format (missing underscore separator): %s", sheetToken)
	}
	
	spreadsheetToken := sheetToken[:lastUnderscore]
	sheetID := sheetToken[lastUnderscore+1:]
	
	// 定义原始 API 响应结构，使用 interface{} 来处理任意类型的值
	type SheetValueResponse struct {
		Code int `json:"code"`
		Msg  string `json:"msg"`
		Data struct {
			ValueRanges []struct {
				MajorDimension string         `json:"majorDimension"`
				Range          string         `json:"range"`
				Values         [][]interface{} `json:"values"`
			} `json:"valueRanges"`
		} `json:"data"`
	}
	
	// 构建请求体
	requestBody := map[string]interface{}{
		"spreadsheetToken": spreadsheetToken,
		"ranges":           []string{sheetID},
	}
	requestJSON, err := json.Marshal(requestBody)
	if err != nil {
		return nil, fmt.Errorf("failed to marshal request: %w", err)
	}
	
	// 创建 HTTP 请求
	// 使用飞书 API 的 endpoint
	url := "https://open.feishu.cn/open-apis/sheets/v4/spreadsheets/" + spreadsheetToken + "/values:batchGet"
	req, err := http.NewRequestWithContext(ctx, "POST", url, bytes.NewBuffer(requestJSON))
	if err != nil {
		return nil, fmt.Errorf("failed to create request: %w", err)
	}
	
	// 设置请求头
	req.Header.Set("Content-Type", "application/json")
	req.Header.Set("Accept", "application/json")
	
	// 获取访问令牌
	// 注意：这里需要从 lark client 获取访问令牌
	// 由于 lark SDK 没有直接提供获取令牌的方法，我们需要使用 SDK 的认证机制
	// 作为一个 workaround，我们使用 SDK 的方法，但手动处理响应
	
	// 尝试使用 SDK 的方法
	valueResp, _, err := c.larkClient.Drive.BatchGetSheetValue(ctx, &lark.BatchGetSheetValueReq{
		SpreadSheetToken: spreadsheetToken,
		Ranges:           []string{sheetID},
	})
	if err != nil {
		// 如果失败，返回详细的错误信息
		return nil, fmt.Errorf("failed to get sheet values: %w", err)
	}

	if len(valueResp.ValueRanges) == 0 {
		return nil, fmt.Errorf("no value ranges found")
	}

	// 转换为二维数组
	values := valueResp.ValueRanges[0].Values
	if len(values) == 0 {
		return nil, fmt.Errorf("sheet is empty")
	}

	// 将 [][]SheetContent 转换为 [][]string
	result := make([][]string, len(values))
	for i, row := range values {
		result[i] = make([]string, len(row))
		for j, cell := range row {
			// 根据单元格类型提取值
			if cell.String != nil {
				result[i][j] = *cell.String
			} else if cell.Int != nil {
				result[i][j] = fmt.Sprintf("%d", *cell.Int)
			} else if cell.Float != nil {
				// 处理浮点数
				result[i][j] = fmt.Sprintf("%g", *cell.Float)
			} else if cell.Link != nil {
				result[i][j] = cell.Link.Text
			} else if cell.Formula != nil {
				// 公式类型，Text 字段存储公式本身
				result[i][j] = cell.Formula.Text
			} else if cell.AtUser != nil {
				result[i][j] = cell.AtUser.Text
			} else if cell.AtDoc != nil {
				result[i][j] = cell.AtDoc.Text
			} else if cell.MultiValue != nil && len(cell.MultiValue.Values) > 0 {
				// 下拉列表，可能有多个值
				values := make([]string, len(cell.MultiValue.Values))
				for k, v := range cell.MultiValue.Values {
					// Values 是 []interface{}，可能是 string, bool, 或 number
					switch val := v.(type) {
					case string:
						values[k] = val
					case bool:
						values[k] = fmt.Sprintf("%t", val)
					case float64:
						values[k] = fmt.Sprintf("%g", val) // 使用 %g 去掉不必要的零
					case int:
						values[k] = fmt.Sprintf("%d", val)
					case int64:
						values[k] = fmt.Sprintf("%d", val)
					default:
						values[k] = fmt.Sprintf("%v", val)
					}
				}
				result[i][j] = strings.Join(values, ", ")
			} else {
				result[i][j] = ""
			}
		}
	}

	return result, nil
}