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
{
  "accident": [
    {
      "description": "title",
      "level": "level",
      "title_merge": "线上事故"
    },
    {
      "description": "title",
      "level": "level",
      "title_merge": "线上事故"
    },
    {
      "description": "title",
      "level": "level",
      "title_merge": "线上事故"
    }
  ],
  "group": "Company Group",
  "manager": "",
  "name": "name-cn-sls-mf-leader",
  "okrs": [
    {
      "category_merge": "Business Project Goals",
      "completion": "",
      "info_merge": "O1: Object1",
      "keyResult": "KR1：what is the name",
      "remark": "",
      "score": 0,
      "weight": 0,
      "weightScore": 0
    },
    {
      "category_merge": "Business Project Goals",
      "completion": "",
      "info_merge": "O1: Object1",
      "keyResult": "KR2：what is the second key result",
      "remark": "",
      "score": 0,
      "weight": 0,
      "weightScore": 0
    },
    {
      "category_merge": "Business Project Goals",
      "completion": "",
      "info_merge": "O2: Object2",
      "keyResult": "KR1：what is the KR1",
      "remark": "",
      "score": 0,
      "weight": 0,
      "weightScore": 0
    },
    {
      "category_merge": "Business Project Goals",
      "completion": "",
      "info_merge": "O2: Object2",
      "keyResult": "KR1：this is okr1",
      "remark": "",
      "score": 0,
      "weight": 0,
      "weightScore": 0
    },
    {
      "category_merge": "Internal Project Goals",
      "completion": "",
      "info_merge": "O3: Object3",
      "keyResult": "KR1: my name is okr1",
      "remark": "",
      "score": 0,
      "weight": 0,
      "weightScore": 0
    },
    {
      "category_merge": "Internal Project Goals",
      "completion": "",
      "info_merge": "O3: Object3",
      "keyResult": "KR2: what is this?",
      "remark": "",
      "score": 0,
      "weight": 0,
      "weightScore": 0
    }
  ],
  "quarter": "Q1",
  "reviewOverall": "",
  "reviewResult": "",
  "role": "Tech Leader",
  "team": "",
  "totalScore": 0
}
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
