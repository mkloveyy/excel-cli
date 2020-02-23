package commands

import (
	"errors"
	"log"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

var companyOwnerMap = map[string]string{
	"珠海横琴极盛科技有限公司":        "李凡",
	"中银国际证券股份有限公司":        "车亮",
	"长江证券股份有限公司":          "邓磊",
	"长江证券（上海）资产管理有限公司":    "关梦姝",
	"长城国瑞证券有限公司":          "廖红飞",
	"野村东方国际证券有限公司":        "马士杰",
	"兴证证券资产管理有限公司":        "李泽江",
	"兴证创新资本管理有限公司":        "李泽江",
	"兴业证券股份有限公司上海证券自营分公司": "张笑燕",
	"兴业证券股份有限公司上海分公司":     "张笑燕",
	"兴业证券股份有限公司":          "张笑燕",
	"星展证券（中国）有限公司（筹）":     "马士杰",
	"湘财证券股份有限公司北京资产管理分公司": "杨庆志",
	"湘财证券股份有限公司":          "杨庆志",
	"西南证券股份有限公司":          "修佃军",
	"西藏东方财富证券股份有限公司":      "胡涛",
	"申万宏源证券有限公司":          "华小惠",
	"申港证券股份有限公司":          "陈文婷",
	"上交所技术有限责任公司":         "陈文婷",
	"上海证券有限责任公司":          "胡涛",
	"上海证券交易所":             "陈文婷",
	"上海海通证券资产管理有限公司":      "杨柳",
	"上海国泰君安证券资产管理有限公司":    "李文玲",
	"上海光大证券资产管理有限公司":      "赵丽",
	"上海东方证券资产管理有限公司":      "陈文婷",
	"南京证券股份有限公司":          "周琳",
	"摩根士丹利华鑫证券有限责任公司":     "何劲松",
	"华鑫证券有限责任公司":          "何劲松",
	"华西证券股份有限公司浙江分公司":     "修佃军",
	"华西证券股份有限公司":          "修佃军",
	"华泰证券股份有限公司南京分公司":     "周琳",
	"华泰证券股份有限公司":          "周琳",
	"华泰证券（上海）资产管理有限公司":    "周琳",
	"华菁证券有限公司":            "胡涛",
	"华金证券股份有限公司":          "李泽江",
	"华福证券有限责任公司":          "廖红飞",
	"华创证券有限责任公司云南分公司":     "张伟",
	"华创证券有限责任公司山东分公司":     "张伟",
	"华创证券有限责任公司":          "张伟",
	"华宝证券有限责任公司":          "华小惠",
	"华安证券股份有限公司":          "陆云峰",
	"宏信证券有限责任公司":          "王琴",
	"海通证券股份有限公司":          "杨柳",
	"国元证券股份有限公司":          "陆云峰",
	"国泰君安证券股份有限公司":        "李文玲",
	"国盛证券资产管理有限公司":        "李凡",
	"国盛证券有限责任公司":          "李凡",
	"国盛期货有限责任公司":          "李凡",
	"国联证券股份有限公司":          "金兵",
	"国金证券股份有限公司上海证券自营分公司": "马士杰",
	"国金证券股份有限公司":          "马士杰",
	"国金道富投资服务有限公司":        "马士杰",
	"光大证券股份有限公司":          "赵丽",
	"方正证券股份有限公司":          "梁筱",
	"东吴证券股份有限公司":          "黄昱辰",
	"东海证券股份有限公司":          "李泽江",
	"东方证券股份有限公司":          "陈文婷",
	"东方花旗证券有限公司":          "陈文婷",
	"灯塔财经信息有限公司":          "杨庆志",
	"德邦证券股份有限公司":          "华小惠",
	"川财证券有限责任公司":          "王琴",
	"财富证券有限责任公司":          "杨庆志",
	"爱建证券有限责任公司":          "胡涛",
}

func Correct(filePath string, inputFileName string, sheetName string, ownerColumn string, companyColumn string) (err error) {
	if len(ownerColumn) > 1 || len(companyColumn) > 1 {
		return errors.New("wrong column value")
	}

	var f *excelize.File

	if f, err = excelize.OpenFile(filePath + inputFileName); err != nil {
		return
	}

	// get sheet
	sRows := f.GetRows(sheetName)
	// get sheet title
	sTitle := sRows[0]

	// output wrong data to a new file
	wrongDataFile := excelize.NewFile()
	wrongDataSheet := wrongDataFile.NewSheet(sheetName)
	// set title
	wrongDataFile.SetSheetRow(sheetName, "A1", &sTitle)
	wrongDataFile.SetActiveSheet(wrongDataSheet)

	// init wrong data count
	count := 0

	// classify by diff column
	for j := 2; j < len(sRows); j++ {
		// get company and pre owner
		company := f.GetCellValue(sheetName, companyColumn+strconv.Itoa(j))
		owner := f.GetCellValue(sheetName, ownerColumn+strconv.Itoa(j))

		for k, v := range companyOwnerMap {
			if k == company && v != owner {
				wrongDataFile.SetSheetRow(sheetName, "A"+strconv.Itoa(count+2), &sRows[j-1])
				log.Printf("row %d wrong: company %s, before is %s, now is %s", j, company, owner, v)
				f.SetCellValue(sheetName, ownerColumn+strconv.Itoa(j), v)
				count++
				break
			}
		}
	}

	log.Printf("total changed: %d rows", count)

	// update file
	if err = f.Save(); err != nil {
		return
	}

	// save wrong data file
	if err = wrongDataFile.SaveAs(filePath + sheetName + "_错误数据" + ".xlsx"); err != nil {
		return
	}

	return
}
