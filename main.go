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
	"unicode"
	"unicode/utf8"
)

func main() {
	xlFile, err := xlsx.OpenFile("input.xlsx")
	if err != nil {
		log.Fatal(err)
	}
	sliseOfSub := make(map[string]string)
	sliseOfRealise := make(map[string]string)
	ostMap := make(map[string]string)
	ostMapUs := make(map[string]string)
	obyazMap := make(map[string]string)
	obyazMapUs := make(map[string]string)
	sheet := xlFile.Sheets[0] // Выберите лист (например, первый лист)
	row := sheet.Rows[6]      // Выберите строку (например, первая строка)
	cell := row.Cells[1]
	rowDate := sheet.Rows[8]
	cellDate := rowDate.Cells[1]
	rowAgree := sheet.Rows[52]
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
				if strings.ToLower(strings.TrimSpace(ndsCellValue)) == "нет" {
					sliseOfRealise[debtValue+" "+strings.ToLower(ndsCellValue)] = value
				} else {
					sliseOfRealise[debtValue] = value
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
				if strings.ToLower(strings.TrimSpace(ndsCellValue)) == "нет" {
					sliseOfSub[debtValue+" "+strings.ToLower(ndsCellValue)] = value
				} else {
					sliseOfSub[debtValue] = value
				}
			} else if columnIndex == 4 {
				cellOst := newRow.Cells[4]
				cellOstValue, err := cellOst.FormattedValue()
				if err != nil {
					continue
				}
				if cellOstValue == "0.00 ₽" || cellOstValue == "" {
				} else {
					cellDoc := newRow.Cells[0]
					cellDocValue, err := cellDoc.FormattedValue()
					if err != nil {
						continue
					}
					if cellDocValue == "" {
						break loop
					}
					ndsCell := newRow.Cells[1]
					ndsCellValue, err := ndsCell.FormattedValue()
					if err != nil {
						log.Println(err)
						continue
					}
					if strings.ToLower(strings.TrimSpace(ndsCellValue)) == "нет" {
						ostMapUs[cellOstValue+" "+strings.ToLower(ndsCellValue)] = cellDocValue
					} else {
						ostMapUs[cellOstValue] = cellDocValue
					}
				}
			} else if columnIndex == 9 {
				cellOst := newRow.Cells[9]
				cellOstValue, err := cellOst.FormattedValue()
				if err != nil {
					continue
				}
				if cellOstValue == "0.00 ₽" || cellOstValue == "" {

				} else {
					cellDoc := newRow.Cells[5]
					cellDocValue, err := cellDoc.FormattedValue()
					if err != nil {
						continue
					}
					if cellDocValue == "" {
						break loop
					}
					ndsCell := newRow.Cells[6]
					ndsCellValue, err := ndsCell.FormattedValue()
					if err != nil {
						log.Println(err)
						continue
					}
					if strings.ToLower(strings.TrimSpace(ndsCellValue)) == "нет" {
						ostMap[cellOstValue+" "+strings.ToLower(ndsCellValue)] = cellDocValue
					} else {
						ostMap[cellOstValue] = cellDocValue
					}
				}
			} else if columnIndex == 3 {
				zachetCell := newRow.Cells[3]
				zachetCellValue, err := zachetCell.FormattedValue()
				if err != nil {
					continue
				}
				if zachetCellValue != "" {
					cellDoc := newRow.Cells[0]
					cellDocValue, err := cellDoc.FormattedValue()
					if err != nil {
						continue
					}
					if cellDocValue == "" {
						break loop
					}
					ndsCell := newRow.Cells[1]
					ndsCellValue, err := ndsCell.FormattedValue()
					if err != nil {
						log.Println(err)
						continue
					}
					if strings.ToLower(strings.TrimSpace(ndsCellValue)) == "нет" {
						obyazMapUs[zachetCellValue+" "+strings.ToLower(ndsCellValue)] = cellDocValue
					} else {
						obyazMapUs[zachetCellValue] = cellDocValue
					}
				} else {
				}
			} else if columnIndex == 8 {
				zachetCell := newRow.Cells[8]
				zachetCellValue, err := zachetCell.FormattedValue()
				if err != nil {
					continue
				}
				if zachetCellValue != "" {
					cellDoc := newRow.Cells[5]
					cellDocValue, err := cellDoc.FormattedValue()
					if err != nil {
						continue
					}
					if cellDocValue == "" {
						break loop
					}
					ndsCell := newRow.Cells[6]
					ndsCellValue, err := ndsCell.FormattedValue()
					if err != nil {
						log.Println(err)
						continue
					}
					if strings.ToLower(strings.TrimSpace(ndsCellValue)) == "нет" {
						obyazMap[zachetCellValue+" "+strings.ToLower(ndsCellValue)] = cellDocValue
					} else {
						obyazMap[zachetCellValue] = cellDocValue
					}
				} else {
				}
			}
		}
	}
	counteragent := cell.String()
	date := cellDate.String()
	t, _ := time.Parse("01-02-06", date)
	dateString := t.Format("02.01.2006")
	agree := cellAgree.String()
	doc := docx.NewA4()
	para2 := doc.AddParagraph().Justification("center")
	para2.AddText("СОГЛАШЕНИЕ")
	para3 := doc.AddParagraph().Justification("center")
	para3.AddText("О ЗАЧЕТЕ ВЗАИМНЫХ ТРЕБОВАНИЙ")
	doc.AddParagraph()
	para := doc.AddParagraph()
	//dateFormat := t.Format("02 января 2006 г.")
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
	para.AddText(" , в лице ______________________________ действующая на основании _______________________________________, с другой стороны,  подписали настоящее Соглашение о нижеследующем:")
	doc.AddParagraph()
	i := 1
	delete(sliseOfRealise, "")
	//Нам должны
	for key, value := range sliseOfRealise {
		value = addY(value)
		forNds, money, isNds, _ := niceType(key)
		nds := forNds * 20 / 120
		ndsString := strconv.FormatFloat(nds, 'f', 2, 64)
		ndsString = strings.ReplaceAll(ndsString, ".", ",")
		rublesText, kopecksText := sumToText(ndsString)
		moneyText, smallMoneyText := sumToText(money)
		moneyText = firstLetterToUpper(moneyText)
		rublesText = firstLetterToUpper(rublesText)
		para = doc.AddParagraph()

		if isNds {
			para.AddText(strconv.Itoa(i) + ". " + counteragent).Bold()
			para.AddText("⠀по состоянию на⠀")
			para.AddText(dateString + "г.").Bold()
			para.AddText("⠀имеет задолженность перед⠀")
			para.AddText("ЗАО «ЭнергоСтрой»").Bold()
			para.AddText("⠀по " + value).Bold()
			para.AddText(". в размере⠀")
			para.AddText(money + "⠀").Bold()
			para.AddText(" руб. (" + moneyText + " " + smallMoneyText + "), в т.ч НДС⠀")
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
			para.AddText(" руб. (" + moneyText + " " + smallMoneyText + "), без учета НДС. Срок исполнения обязательств наступил.")
			para2.AddTab()
			para2 = doc.AddParagraph()
			para2.AddText("Наличие указанной задолженности подтверждается Актом сверки взаиморасчетов по состоянию на " + dateString + "г.")
		}
		i++
	}
	delete(sliseOfSub, "")
	//Мы должны
	for key, value := range sliseOfSub {
		value = addY(value)
		forNds, money, isNds, _ := niceType(key)
		nds := forNds * 20 / 120
		ndsString := strconv.FormatFloat(nds, 'f', 2, 64)
		ndsString = strings.ReplaceAll(ndsString, ".", ",")
		rublesText, kopecksText := sumToText(ndsString)
		moneyText, smallMoneyText := sumToText(money)
		moneyText = firstLetterToUpper(moneyText)
		rublesText = firstLetterToUpper(rublesText)
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
			para.AddText(" руб. (" + moneyText + " " + smallMoneyText + "), в т.ч НДС⠀")
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
			para.AddText(" руб. (" + moneyText + " " + smallMoneyText + "), без учета НДС. Срок исполнения обязательств наступил.")
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
	para.AddText("Стороны пришли к соглашению о зачете взаимных  требований в соответствии со ст. 410 ГК РФ по обязательствам, указанным в п. 1 - " + strconv.Itoa(i-1) + " настоящего соглашения, в размере⠀")
	para.AddText(agreeNdsText).Bold()
	para.AddText("⠀ руб. (" + moneyText + " " + smallMoneyText + "), в т.ч. НДС⠀")
	para.AddText(ndsAString + "⠀").Bold()
	para.AddText(" руб. (" + rublesText + " " + kopecksText + " " + ")")
	para = doc.AddParagraph()
	b := 1
	//Обязательства нам должны
	for key, value := range obyazMap {
		value = addY(value)
		forNds, money, isNds, _ := niceType(key)
		nds := forNds * 20 / 120
		ndsString := strconv.FormatFloat(nds, 'f', 2, 64)
		ndsString = strings.ReplaceAll(ndsString, ".", ",")
		rublesText, kopecksText := sumToText(ndsString)
		moneyText, smallMoneyText := sumToText(money)
		moneyText = firstLetterToUpper(moneyText)
		rublesText = firstLetterToUpper(rublesText)
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
			para.AddText(" руб. (" + moneyText + " " + smallMoneyText + "), в т.ч НДС⠀")
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
			para.AddText(" руб. (" + moneyText + " " + smallMoneyText + "), без учета НДС. с " + dateString + "г.")
		}
		b++
	}
	//Обязательства мы должны
	for key, value := range obyazMapUs {
		value = addY(value)
		forNds, money, isNds, _ := niceType(key)
		nds := forNds * 20 / 120
		ndsString := strconv.FormatFloat(nds, 'f', 2, 64)
		ndsString = strings.ReplaceAll(ndsString, ".", ",")
		rublesText, kopecksText := sumToText(ndsString)
		moneyText, smallMoneyText := sumToText(money)
		moneyText = firstLetterToUpper(moneyText)
		rublesText = firstLetterToUpper(rublesText)
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
			para.AddText(" руб. (" + moneyText + " " + smallMoneyText + "), в т.ч НДС⠀")
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
			para.AddText(" руб. (" + moneyText + " " + smallMoneyText + "), без учета НДС. с " + dateString + "г.")
		}
		b++
	}
	para = doc.AddParagraph()
	i++
	para.AddText(strconv.Itoa(i) + ".⠀").Bold()
	para.AddText("С момента подписания настоящего Соглашения стороны считают себя свободными от обязательств, в размере, прекращенном зачетом согласно п.10 настоящего соглашения.")
	para = doc.AddParagraph()
	para = doc.AddParagraph()
	para.AddText("Настоящее Соглашение составлено в 2-х подлинных экземплярах, по одному для каждой из сторон.")
	para = doc.AddParagraph()
	para.AddText("Настоящее Соглашение вступает в силу с момента его подписания сторонами.")
	para = doc.AddParagraph()
	para = doc.AddParagraph()
	//Нам должны
	for key, value := range ostMap {
		value = addY(value)
		forNds, money, isNds, _ := niceType(key)
		nds := forNds * 20 / 120
		ndsString := strconv.FormatFloat(nds, 'f', 2, 64)
		ndsString = strings.ReplaceAll(ndsString, ".", ",")
		rublesText, kopecksText := sumToText(ndsString)
		moneyText, smallMoneyText := sumToText(money)
		moneyText = firstLetterToUpper(moneyText)
		rublesText = firstLetterToUpper(rublesText)
		if isNds {
			para.AddText(strconv.Itoa(i) + ".⠀").Bold()
			para.AddText("После проведения соглашение сохраняется задолженность в пользу⠀")
			para.AddText("ЗАО «ЭнергоСтрой»⠀").Bold()
			para.AddText("по " + value + ". в размере⠀")
			para.AddText(money).Bold()
			para.AddText(" руб. (" + moneyText + " " + smallMoneyText + ", в т.ч НДС⠀")
			para.AddText(ndsString + "⠀").Bold()
			para.AddText(" руб. (" + rublesText + " " + kopecksText + ")")
		} else {
			para.AddText(strconv.Itoa(i) + ".⠀").Bold()
			para.AddText("После проведения соглашение сохраняется задолженность в пользу⠀")
			para.AddText("ЗАО «ЭнергоСтрой»⠀").Bold()
			para.AddText("по " + value + ". в размере⠀")
			para.AddText(money).Bold()
			para.AddText(" руб. (" + moneyText + " " + smallMoneyText + "), без НДС.")
		}
		i++
		para = doc.AddParagraph()
	}
	// Мы должны
	for key, value := range ostMapUs {
		value = addY(value)
		forNds, money, isNds, _ := niceType(key)
		nds := forNds * 20 / 120
		ndsString := strconv.FormatFloat(nds, 'f', 2, 64)
		ndsString = strings.ReplaceAll(ndsString, ".", ",")
		rublesText, kopecksText := sumToText(ndsString)
		moneyText, smallMoneyText := sumToText(money)
		moneyText = firstLetterToUpper(moneyText)
		rublesText = firstLetterToUpper(rublesText)
		if isNds {
			para.AddText(strconv.Itoa(i) + ".⠀").Bold()
			para.AddText("После проведения соглашение сохраняется задолженность в пользу⠀")
			para.AddText(counteragent).Bold()
			para.AddText("⠀по " + value + ". в размере⠀")
			para.AddText(money).Bold()
			para.AddText(" руб. (" + moneyText + " " + smallMoneyText + ", в т.ч НДС⠀")
			para.AddText(ndsString + "⠀").Bold()
			para.AddText(" руб. (" + rublesText + " " + kopecksText + ")")
		} else {
			para.AddText(strconv.Itoa(i) + ".⠀").Bold()
			para.AddText("После проведения соглашение сохраняется задолженность в пользу⠀")
			para.AddText(counteragent).Bold()
			para.AddText("⠀по " + value + ". в размере⠀")
			para.AddText(money).Bold()
			para.AddText(" руб. (" + moneyText + " " + smallMoneyText + "), без НДС.")
		}
		i++
		para = doc.AddParagraph()
	}

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
func firstLetterToUpper(s string) string {
	if len(s) > 0 {
		r, size := utf8.DecodeRuneInString(s)
		if r != utf8.RuneError || size > 1 {
			lower := unicode.ToUpper(r)
			if lower != r {
				s = string(lower) + s[size:]
			}
		}
	}
	return s
}
func sumToText(text string) (string, string) {
	parts := strings.Split(text, ",")
	rublesPart := parts[0]
	kopecksPart := parts[1]
	rublesText := numtow.MustString(rublesPart, lang.RU) + " рублей"
	kopecksText := kopecksPart + " копеек"
	return rublesText, kopecksText
}
func niceType(amountStr string) (float64, string, bool, error) {
	// Убираем символ валюты и разделитель тысяч (если есть)
	isNds := true
	amountStr = strings.ReplaceAll(amountStr, " ₽", "")
	parts := strings.Split(amountStr, " ")
	if len(parts) > 1 {
		if parts[1] == "нет" {
			isNds = false
		}
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
