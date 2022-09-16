package main

import (
	"encoding/json"
	xlst "github.com/sjqzhang/go-xlsx-templater"
	"net/http"
	"path/filepath"

	"github.com/gin-gonic/gin"
)

func main() {
	router := gin.Default()
	// Set a lower memory limit for multipart forms (default is 32 MiB)
	router.MaxMultipartMemory = 8 << 20 // 8 MiB
	router.GET("/", func(context *gin.Context) {
		html := `
<!doctype html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <title>Upload Excel Template For Render</title>
</head>
<body>
<h1>Upload Excel Template For Render</h1>

<form action="/render" method="post" enctype="multipart/form-data">
    <p>
    JSON Data: <textarea rows="20" cols="50" name="data" name="name">
{"accident":[{"description":"title","level":"level"},{"description":"title","level":"level"},{"description":"title","level":"level"}],"group":"Supply Chain","manager":"","name":"name-cn-sls-mf-leader","okrs":[{"category":"Business Project Goals","completion":"","info":"Data Mart: SLS Mart\nAccomplish TWS module for sls mart, guarante 80% of OFM users migrate to sls mart.","keyResult":"KR1：Start construction of TWS Mart, go UAT at the end of June.","remark":"","score":0,"weight":0,"weightScore":0},{"category":"Business Project Goals","completion":"","info":"Data Mart: SLS Mart\nAccomplish TWS module for sls mart, guarante 80% of OFM users migrate to sls mart.","keyResult":"KR2：At the end of June, the migration of local users of OFM and users of SLS union tables and marketplace OFM will be completed.","remark":"","score":0,"weight":0,"weightScore":0},{"category":"Business Project Goals","completion":"","info":"Data Product：Mart Configuration Tool\nAccomplish Mart Configuration Tool, migrate some of BI users' offline source table with MCT ","keyResult":"KR1：End of April, Mart Configuration Tool phase 1 startup，complete technical design","remark":"","score":0,"weight":0,"weightScore":0},{"category":"Business Project Goals","completion":"","info":"Data Product：Mart Configuration Tool\nAccomplish Mart Configuration Tool, migrate some of BI users' offline source table with MCT ","keyResult":"KR1：End of May, Mart Configuration Tool phase 1 - Go UAT","remark":"","score":0,"weight":0,"weightScore":0},{"category":"Internal Project Goals","completion":"","info":"Data Quality Improvement\nBuild automatic test procedure and test case lib to improve efficiency","keyResult":"KR1: Mid of May, build up test case lib for all domains","remark":"","score":0,"weight":0,"weightScore":0},{"category":"Internal Project Goals","completion":"","info":"Data Quality Improvement\nBuild automatic test procedure and test case lib to improve efficiency","keyResult":"KR2: End of Apr, integrate test cases with Athena, generate automatic execution tasks.","remark":"","score":0,"weight":0,"weightScore":0}],"quarter":"Q1","reviewOverall":"","reviewResult":"","role":"Tech Leader","team":"","totalScore":0}
</textarea>
   </p>
    <p>
    Files: <input type="file"  name="file">
  </p>
    <input type="submit" value="Submit">
</form>
</body>
`
		context.Header("Content-Type", "text/html; charset=utf-8")
		context.String(http.StatusOK, html)
	})
	router.POST("/render", func(c *gin.Context) {
		data := c.PostForm("data")
		file, err := c.FormFile("file")
		if err != nil {
			c.String(http.StatusBadRequest, "get form err: %s", err.Error())
			return
		}
		filename := filepath.Base(file.Filename)
		if err := c.SaveUploadedFile(file, filename); err != nil {
			c.String(http.StatusBadRequest, "upload file err: %s", err.Error())
			return
		}
		doc := xlst.New()
		err = doc.ReadTemplate(file.Filename)
		if err != nil {
			c.String(http.StatusBadRequest, "read template err: %s", err.Error())
			return
		}
		var ctx map[string]interface{}
		err = json.Unmarshal([]byte(data), &ctx)
		if err != nil {
			c.String(http.StatusBadRequest, "data Unmarshal error: %s", err.Error())
			return
		}
		err = doc.Render(ctx)
		if err != nil {
			c.String(http.StatusBadRequest, "render error: %s", err.Error())
			return
		}
		c.Header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
		doc.Write(c.Writer)
	})
	router.Run(":8080")
}
