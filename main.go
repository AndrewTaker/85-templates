package main

import (
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

// Struct that encapsulates data about cell.
//
// row:
//
// integer, literal value of the worksheet row.
//
// formula:
//
// template formula. Exists solely for future replacement inside VLOOKUP formula for specific worksheet.
type cellData struct {
	row     int
	formula string
}

const (
	fileName         string = "test.xlsx" // Name of the file containing pivot table and template data.
	defaultSheetName string = "template"  // Name of the worksheet containing template itself.
	maxRow           int    = 2074        // Max rows of template worksheet. It is better to make manual input rather than counting last row.
	pivotWorksheet   string = "s"         // Name of the worksheet containig pivot table.
	hidenWorksheet   string = "valid"     // Optional: name of the hidden worksheet.
)

// Creates N copies of template worksheet. Mostly for testing purposes.
func createDummyWorksheets(f *excelize.File, times int) {
	counter := 1
	templateIndex, _ := f.GetSheetIndex(defaultSheetName)
	for counter < times {
		newSheetIndex, _ := f.NewSheet(fmt.Sprintf("доо%v", counter))
		err := f.CopySheet(templateIndex, newSheetIndex)
		if err != nil {
			fmt.Println("Could not create worksheet", err)
			return
		}
		counter += 1
	}
}

// Copies template worksheet with given names.
func copyTemplate(f *excelize.File, sheetNames []string) {
	templateIndex, _ := f.GetSheetIndex(defaultSheetName)
	for _, worksheet := range sheetNames {
		newSheetIndex, _ := f.NewSheet(worksheet)
		err := f.CopySheet(templateIndex, newSheetIndex)
		if err != nil {
			fmt.Println(err)
			return
		}
	}
}

// Creates a slice with template formulas.
//
// Default values based on current template workbook: rowIndex = 1, columnIndex = 8
func createSliceWithTemplateFormulas(f *excelize.File, rowIndex int, columnIndex int) []*cellData {
	var data = []*cellData{}
	for rowIndex < maxRow {
		address, _ := excelize.CoordinatesToCellName(columnIndex, rowIndex)
		value, _ := f.GetCellValue(pivotWorksheet, address)
		formula, _ := f.GetCellFormula(pivotWorksheet, address)
		encapsulatedData := &cellData{
			row:     rowIndex,
			formula: formula,
		}
		data = append(data, encapsulatedData)
		fmt.Printf("ADDRESS: %v\tVALUE: %v\tFORMULA: %v\n", address, value, formula)
		rowIndex += 1
	}
	return data
}

// Copies formulas from template main row to adjacent rows.
// Default values based on current template workbook: columnIndex = 9
func copyFormulasFromTemplateColumn(f *excelize.File, data []*cellData, columnIndex int) {
	worksheets := f.GetSheetList()
	for _, worksheet := range worksheets {
		if worksheet != pivotWorksheet && worksheet != defaultSheetName && worksheet != hidenWorksheet {
			fmt.Printf("WORKING ON %s\n", worksheet)
			for _, cellData := range data {
				columnLetter, _ := excelize.ColumnNumberToName(columnIndex)
				cellAddress := fmt.Sprintf("%s%v", columnLetter, cellData.row)
				err := f.SetCellFormula(pivotWorksheet, cellAddress, strings.Replace(cellData.formula, defaultSheetName, worksheet, -1))
				f.SetCellValue(pivotWorksheet, fmt.Sprintf("%s2", columnLetter), worksheet)
				if err != nil {
					fmt.Println(err)
					return
				}
			}
			columnIndex += 1
		}
	}
}

func main() {
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		fmt.Println(err)
		return
	}
	// Deferring value updating to make sure formulas are correctly displaying VLOOKUP values.
	defer f.UpdateLinkedValue()
	// Deferring workbook closing to make sure it cleans up cache and closes the book after stacks empties.
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// Create slice of kindergartens.
	var kinderGartens = []string{"kg1", "kg2", "kg3", "kg4", "kg5"}
	// Make copies of prepared template worksheet and make new worksheets based on kindergartens slice.
	copyTemplate(f, kinderGartens)
	// Creates slice with struct that contains row number and formula.
	data := createSliceWithTemplateFormulas(f, 1, 8)
	// Creates formulas for each new worksheet on the pivot table worksheet.
	copyFormulasFromTemplateColumn(f, data, 9)
	f.Save()
	fmt.Println("Done.")
}
