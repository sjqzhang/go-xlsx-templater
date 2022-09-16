package main

import (
	"encoding/json"
	xlst "github.com/sjqzhang/go-xlsx-templater"
)

func main() {
	doc := xlst.New()
	doc.ReadTemplate("./demo/okrTemplate.xlsx")
	var ctx map[string]interface{}
	json.Unmarshal([]byte(js), &ctx)
	doc.Render(ctx)
	doc.Save("./demo/report.xlsx")
}

//
//func main() {
//	doc := xlst.New()
//	doc.ReadTemplate("./template.xlsx")
//  ctx := map[string]interface{}{
//        "name": "Github User",
//        "groupHeader": "Group name",
//        "nameHeader": "Item name",
//        "quantityHeader": "Quantity",
//        "groups": []map[string]interface{}{
//            {
//                "name":  "Work",
//                "total": 3,
//                "items": []map[string]interface{}{
//                    {
//                        "name":     "Pen",
//                        "quantity": 2,
//                    },
//                    {
//                        "name":     "Pencil",
//                        "quantity": 1,
//                    },
//                },
//            },
//            {
//                "name":  "Weekend",
//                "total": 36,
//                "items": []map[string]interface{}{
//                    {
//                        "name":     "Condom",
//                        "quantity": 12,
//                    },
//                    {
//                        "name":     "Beer",
//                        "quantity": 24,
//                    },
//                },
//            },
//        },
//    }
//	err := doc.Render(ctx)
//	if err != nil {
//		panic(err)
//	}
//	err = doc.Save("./report.xlsx")
//	if err != nil {
//		panic(err)
//	}
//}

const js = `
{
    "accident": [
        {
            "description": "title",
            "level": "level",
             "title_merge":"线上事故"
        },
        {
            "description": "title",
            "level": "level",
             "title_merge":"线上事故"
        },
        {
            "description": "title",
            "level": "level",
             "title_merge":"线上事故"
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
            "info_merge":"O2: Object2",
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

`
