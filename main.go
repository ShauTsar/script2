package main

import (
	"fmt"
	"github.com/fumiama/go-docx"
	"github.com/gammban/numtow"
	"github.com/gammban/numtow/lang"
	"github.com/tealeg/xlsx"
	"log"
	"os"
	"strconv"
	"strings"
	"time"
)

func main() {
	xlFile, err := xlsx.OpenFile("input.xlsx")
	if err != nil {
		log.Fatal(err)
	}
	sliseOfSub := make(map[string]string)
	sliseOfRealise := make(map[string]string)
	sheet := xlFile.Sheets[0] // Выберите лист (например, первый лист)
	row := sheet.Rows[6]      // Выберите строку (например, первая строка)
	cell := row.Cells[1]
	rowDate := sheet.Rows[8]
	cellDate := rowDate.Cells[1]
	rowAgree := sheet.Rows[51]
	cellAgree := rowAgree.Cells[2]
	startRowRealise := 13
loop:
	for rowIndex := startRowRealise; rowIndex <= sheet.MaxRow; rowIndex++ {
		newRow := sheet.Rows[rowIndex]
		for columnIndex, newCell := range newRow.Cells {
			if columnIndex == 5 {
				value, err := newCell.FormattedValue()
				if err != nil {
					log.Println(err)
					continue
				}
				if value == "Итого по договорам:" {
					break loop

				}
				debt := newRow.Cells[7]
				debtValue, err := debt.FormattedValue()
				if err != nil {
					log.Println(err)
					continue
				}
				ndsCell := newRow.Cells[6]
				ndsCellValue, err := ndsCell.FormattedValue()
				if err != nil {
					log.Println(err)
					continue
				}
				if strings.ToLower(ndsCellValue) == "нет" && value != "" {
					sliseOfRealise[debtValue+" "+strings.ToLower(ndsCellValue)] = value
				} else {
					sliseOfSub[debtValue] = value
				}

			} else if columnIndex == 0 {
				value, err := newCell.FormattedValue()
				if err != nil {
					//log.Println(err)
					continue
				}
				if value == "Итого по договорам:" {
					break loop

				}
				debt := newRow.Cells[2]
				debtValue, err := debt.FormattedValue()
				if err != nil {
					//log.Println(err)
					continue
				}
				ndsCell := newRow.Cells[1]
				ndsCellValue, err := ndsCell.FormattedValue()
				if err != nil {
					log.Println(err)
					continue
				}
				if strings.ToLower(ndsCellValue) == "нет" && value != "" {
					sliseOfRealise[debtValue+" "+strings.ToLower(ndsCellValue)] = value
				} else {
					sliseOfSub[debtValue] = value
				}
			}
		}
	}
	counteragent := cell.String()
	date := cellDate.String()
	agree := cellAgree.String()
	doc := docx.NewA4()
	para2 := doc.AddParagraph().Justification("center")
	para2.AddText("СОГЛАШЕНИЕ")
	para3 := doc.AddParagraph().Justification("center")
	para3.AddText("О ЗАЧЕТЕ ВЗАИМНЫХ ТРЕБОВАНИЙ")
	doc.AddParagraph()
	para := doc.AddParagraph()
	para.AddText(fmt.Sprintf("г. Екатеринбург                                                                                                                      \"%s\"", "дата"))
	para = doc.AddParagraph()
	para = doc.AddParagraph()
	para = doc.AddParagraph()
	para.AddTab()
	para.AddText("ЗАО «ЭнергоСтрой»").Bold()
	para.AddText(", в лице Генерального директора Бурнева Б.В., действующего на основании Устава, с одной стороны, и ")
	para = doc.AddParagraph()
	para.AddTab()
	para.AddText(counteragent).Bold()
	para.AddText(" , действующая на основании ____, с другой стороны,  подписали настоящее Соглашение о нижеследующем:")
	doc.AddParagraph()
	i := 1
	t, _ := time.Parse("01-02-06", date)
	dateString := t.Format("02.01.2006")
	delete(sliseOfRealise, "")
	for key, value := range sliseOfRealise {
		value = addY(value)
		forNds, money, isNds, _ := niceType(key)
		nds := forNds * 20 / 120
		ndsString := strconv.FormatFloat(nds, 'f', 2, 64)
		ndsString = strings.ReplaceAll(ndsString, ".", ",")
		rublesText, kopecksText := sumToText(ndsString)
		moneyText, smallMoneyText := sumToText(money)
		para = doc.AddParagraph()

		if isNds {
			//⠀вместо пробела(⠀)
			para.AddText(strconv.Itoa(i) + ". " + counteragent).Bold()
			para.AddText("⠀по состоянию на⠀")
			para.AddText(dateString + "г.").Bold()
			para.AddText("⠀имеет задолженность перед⠀")
			para.AddText("ЗАО «ЭнергоСтрой»").Bold()
			para.AddText("⠀по " + value).Bold()
			para.AddText(". в размере⠀")
			para.AddText(money + "⠀").Bold()
			para.AddText(" руб .(" + moneyText + " " + smallMoneyText + "),в т.ч НДС⠀")
			para.AddText(ndsString + "⠀").Bold()
			para.AddText(" руб. (" + rublesText + " " + kopecksText + "). Срок исполнения обязательств наступил.")
			para2.AddTab()
			para2 = doc.AddParagraph()
			para2.AddText("Наличие указанной задолженности подтверждается Актом сверки взаиморасчетов по состоянию на " + dateString + "г.")

		} else {
			para.AddText(strconv.Itoa(i) + ". " + counteragent).Bold()
			para.AddText("⠀по состоянию на⠀")
			para.AddText(dateString + "г.").Bold()
			para.AddText("⠀имеет задолженность перед⠀")
			para.AddText("ЗАО «ЭнергоСтрой»").Bold()
			para.AddText("⠀по " + value).Bold()
			para.AddText(". в размере⠀")
			para.AddText(money + "⠀").Bold()
			para.AddText(" руб .(" + moneyText + " " + smallMoneyText + "), без учета НДС. Срок исполнения обязательств наступил.")
			para2.AddTab()
			para2 = doc.AddParagraph()
			para2.AddText("Наличие указанной задолженности подтверждается Актом сверки взаиморасчетов по состоянию на " + dateString + "г.")
		}
		i++
	}
	delete(sliseOfSub, "")
	for key, value := range sliseOfSub {
		value = addY(value)
		forNds, money, isNds, _ := niceType(key)
		nds := forNds * 20 / 120
		ndsString := strconv.FormatFloat(nds, 'f', 2, 64)
		ndsString = strings.ReplaceAll(ndsString, ".", ",")
		rublesText, kopecksText := sumToText(ndsString)
		moneyText, smallMoneyText := sumToText(money)
		para = doc.AddParagraph()

		if isNds {
			para.AddText(strconv.Itoa(i) + ". ЗАО «ЭнергоСтрой»").Bold()
			para.AddText("⠀по состоянию на⠀")
			para.AddText(dateString + "г.").Bold()
			para.AddText("⠀имеет задолженность перед⠀")
			para.AddText(counteragent).Bold()
			para.AddText("⠀по⠀")
			para.AddText(value).Bold()
			para.AddText(". в размере⠀")
			para.AddText(money + "⠀").Bold()
			para.AddText(" руб .(" + moneyText + " " + smallMoneyText + "),в т.ч НДС⠀")
			para.AddText(ndsString + "⠀").Bold()
			para.AddText(" руб. (" + rublesText + " " + kopecksText + "). Срок исполнения обязательств наступил.")
			para2.AddTab()
			para2 = doc.AddParagraph()
			para2.AddText("Наличие указанной задолженности подтверждается Актом сверки взаиморасчетов по состоянию на " + dateString + "г.")
		} else {
			para.AddText(strconv.Itoa(i) + ". ЗАО «ЭнергоСтрой»").Bold()
			para.AddText("⠀по состоянию на⠀")
			para.AddText(dateString + "г.").Bold()
			para.AddText("⠀имеет задолженность перед⠀")
			para.AddText(counteragent).Bold()
			para.AddText("⠀по⠀")
			para.AddText(value).Bold()
			para.AddText(". в размере⠀")
			para.AddText(money + "⠀").Bold()
			para.AddText(" руб .(" + moneyText + " " + smallMoneyText + "), без учета НДС. Срок исполнения обязательств наступил.")
			para2 = doc.AddParagraph()
			para2.AddTab()
			para2.AddText("Наличие указанной задолженности подтверждается Актом сверки взаиморасчетов по состоянию на " + dateString + "г.")
		}

		i++
	}
	para = doc.AddParagraph()
	agreeNdsX, agreeNdsText, _, _ := niceType(agree)
	ndsA := agreeNdsX * 20 / 120
	ndsAString := strconv.FormatFloat(ndsA, 'f', 2, 64)
	ndsAString = strings.ReplaceAll(ndsAString, ".", ",")
	rublesText, kopecksText := sumToText(ndsAString)
	moneyText, smallMoneyText := sumToText(agreeNdsText)
	para.AddText(strconv.Itoa(i) + ". ").Bold()
	para.AddText("Стороны пришли к соглашению о зачете взаимных  требований в соответствии со ст. 410 ГК РФ по обязательствам, указанным в п. 1 - 9 настоящего соглашения, в размере⠀")
	para.AddText(agreeNdsText).Bold()
	para.AddText("⠀ руб. (" + moneyText + " " + smallMoneyText + "), в т.ч. НДС⠀")
	para.AddText(ndsAString + "⠀").Bold()
	para.AddText(" руб.(" + rublesText + " " + kopecksText + " " + ")")
	para = doc.AddParagraph()
	b := 1
	for key, value := range sliseOfRealise {
		value = addY(value)
		forNds, money, isNds, _ := niceType(key)
		nds := forNds * 20 / 120
		ndsString := strconv.FormatFloat(nds, 'f', 2, 64)
		ndsString = strings.ReplaceAll(ndsString, ".", ",")
		rublesText, kopecksText := sumToText(ndsString)
		moneyText, smallMoneyText := sumToText(money)
		para = doc.AddParagraph()
		if isNds {
			//⠀вместо пробела(⠀)
			para.AddText(strconv.Itoa(i) + "." + strconv.Itoa(b)).Bold()
			para.AddText(". Обязательства⠀")
			para.AddText(counteragent).Bold()
			para.AddText("⠀перед⠀")
			para.AddText("ЗАО «ЭнергоСтрой»").Bold()
			para.AddText("⠀по⠀")
			para.AddText(value).Bold()
			para.AddText(". прекращаются в размере⠀")
			para.AddText(money + "⠀").Bold()
			para.AddText(" руб .(" + moneyText + " " + smallMoneyText + "),в т.ч НДС⠀")
			para.AddText(ndsString + "⠀").Bold()
			para.AddText(" руб. (" + rublesText + " " + kopecksText + ")с " + dateString + "г.")

		} else {
			para.AddText(strconv.Itoa(i) + "." + strconv.Itoa(b)).Bold()
			para.AddText(". Обязательства⠀")
			para.AddText(counteragent).Bold()
			para.AddText("⠀перед⠀")
			para.AddText("ЗАО «ЭнергоСтрой»").Bold()
			para.AddText("⠀по⠀")
			para.AddText(value).Bold()
			para.AddText(". прекращаются в размере⠀")
			para.AddText(money + "⠀").Bold()
			para.AddText(" руб .(" + moneyText + " " + smallMoneyText + "), без учета НДС. с " + dateString + "г.")
		}
		b++
	}
	for key, value := range sliseOfSub {
		value = addY(value)
		forNds, money, isNds, _ := niceType(key)
		nds := forNds * 20 / 120
		ndsString := strconv.FormatFloat(nds, 'f', 2, 64)
		ndsString = strings.ReplaceAll(ndsString, ".", ",")
		rublesText, kopecksText := sumToText(ndsString)
		moneyText, smallMoneyText := sumToText(money)
		para = doc.AddParagraph()
		if isNds {
			para.AddText(strconv.Itoa(i) + "." + strconv.Itoa(b)).Bold()
			para.AddText(". Обязательства⠀")
			para.AddText("ЗАО «ЭнергоСтрой»").Bold()
			para.AddText("⠀перед⠀")
			para.AddText(counteragent).Bold()
			para.AddText("⠀по ")
			para.AddText(value).Bold()
			para.AddText(". прекращаются в размере⠀")
			para.AddText(money + "⠀").Bold()
			para.AddText(" руб .(" + moneyText + " " + smallMoneyText + "),в т.ч НДС⠀")
			para.AddText(ndsString + "⠀").Bold()
			para.AddText(" руб. (" + rublesText + " " + kopecksText + ")с " + dateString + "г.")
		} else {
			para.AddText(strconv.Itoa(i) + "." + strconv.Itoa(b)).Bold()
			para.AddText(". Обязательства⠀")
			para.AddText("ЗАО «ЭнергоСтрой»").Bold()
			para.AddText("⠀перед⠀")
			para.AddText(counteragent).Bold()
			para.AddText("⠀по ")
			para.AddText(value).Bold()
			para.AddText(". прекращаются в размере⠀")
			para.AddText(money + "⠀").Bold()
			para.AddText(" руб .(" + moneyText + " " + smallMoneyText + "), без учета НДС. с " + dateString + "г.")
		}
		b++
	}
	para = doc.AddParagraph()
	para.AddText(strconv.Itoa(i) + ".⠀").Bold()
	para.AddText("С момента подписания настоящего Соглашения стороны считают себя свободными от обязательств, в размере, прекращенном зачетом согласно п.10 настоящего соглашения.")
	para = doc.AddParagraph()
	para = doc.AddParagraph()
	para.AddText("Настоящее Соглашение составлено в 2-х подлинных экземплярах, по одному для каждой из сторон.")
	para = doc.AddParagraph()
	para.AddText("Настоящее Соглашение вступает в силу с момента его подписания сторонами.")
	para = doc.AddParagraph()
	para = doc.AddParagraph()
	para.AddText(strconv.Itoa(i) + ".⠀").Bold()
	para.AddText("После проведения соглашения сохраняется задолженность в пользу  ____  по договору ____ от ___г. в размере ___, без НДС.")
	para = doc.AddParagraph()
	para = doc.AddParagraph()
	para = doc.AddParagraph()
	para.AddText("ЗАО «ЭнергоСтрой»").Bold().AddTab().AddTab().AddTab().AddTab()
	para.AddText(counteragent).Bold()
	para = doc.AddParagraph()
	para.AddText("Генеральный директор")
	para = doc.AddParagraph()
	para.AddText("_________________ / Б.В. Бурнев /").AddTab().AddTab()
	para.AddText("_________________ / /")
	para = doc.AddParagraph()
	para.AddText("Главный бухгалтер")
	para = doc.AddParagraph()
	para.AddText("_________________ / С.М. Соколова /")
	fileNameDate := time.Now().Format("02.01.2006")
	f, err := os.Create(fileNameDate + ".docx")
	// save to file
	if err != nil {
		panic(err)
	}
	_, err = doc.WriteTo(f)
	if err != nil {
		panic(err)
	}
	err = f.Close()
	if err != nil {
		panic(err)
	}
}
func sumToText(text string) (string, string) {
	parts := strings.Split(text, ",")
	rublesPart := parts[0]
	kopecksPart := parts[1]
	rublesText := numtow.MustString(rublesPart, lang.RU) + " рублей"
	kopecksText := numtow.MustString(kopecksPart, lang.RU) + " копеек"
	return rublesText, kopecksText
}
func niceType(amountStr string) (float64, string, bool, error) {
	// Убираем символ валюты и разделитель тысяч (если есть)
	isNds := true
	amountStr = strings.ReplaceAll(amountStr, "₽", "")
	parts := strings.Split(amountStr, " ")
	if parts[1] == "нет" {
		isNds = false
	}
	amountStr = strings.ReplaceAll(amountStr, "нет", "")
	amountStr = strings.ReplaceAll(amountStr, " ", "")

	// Преобразовываем строку в число с плавающей точкой
	amount, err := strconv.ParseFloat(amountStr, 64)
	if err != nil {
		fmt.Println("Ошибка разбора суммы:", err)
		return 0, "", isNds, err
	}
	formattedAmount := strconv.FormatFloat(amount, 'f', 2, 64)
	formattedAmount = strings.ReplaceAll(formattedAmount, ".", ",")
	return amount, formattedAmount, isNds, nil
}
func addY(text string) string {
	text = strings.ToLower(text)
	parts := strings.Split(text, " ")

	// Ищем слово "договор" и добавляем "у" к нему
	for i, part := range parts {
		if part == "договор" {
			parts[i] = part + "у"
		} else if part == "года" {
			parts[i] = "г"
		}
	}

	// Объединяем подстроки обратно в одну строку
	text = strings.Join(parts, " ")

	return text
}
