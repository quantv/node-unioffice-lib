package main

/*
#include <stdint.h>
#include <errno.h>
typedef uint64_t Handle;

typedef enum {
  ss_ok,
  ss_worksheet_error,
  ss_save_failed,
  ss_not_a_number,
} ss_status;

typedef struct {
	char* v;
	uint8_t t;
	char* s;
} cellValue;
*/
import "C"

import (
	"fmt"
	"time"
	"unsafe"

	"github.com/unidoc/unioffice/spreadsheet"
	"github.com/unidoc/unioffice/spreadsheet/convert2"
)

//export ss_new
func ss_new() C.Handle {
	ss := spreadsheet.New()
	handle := C.Handle(uintptr(unsafe.Pointer(&ss)))
	workbooks[handle] = ss
	return handle
}

//export ss_open
func ss_open(filepath *C.char) C.Handle {
	file := C.GoString(filepath)
	ss, err := spreadsheet.Open(file)
	if err != nil {
		return 0
	}
	handle := C.Handle(uintptr(unsafe.Pointer(&ss)))
	workbooks[handle] = ss
	return handle
}

//export ss_add_sheet
func ss_add_sheet(h C.Handle) *C.char {
	ss := workbooks[h]
	sh := ss.AddSheet()
	return C.CString(sh.Name())
}

//export ss_add_row
func ss_add_row(h C.Handle, sheet *C.char, row *C.uint32_t) C.ss_status {
	sh, err := get_sheet(h, sheet)
	if err != nil {
		return C.ss_worksheet_error
	}
	r := sh.AddRow()
	*row = C.uint32_t(r.RowNumber())

	return C.ss_ok
}

//export ss_add_rows
func ss_add_rows(h C.Handle, sheet *C.char, count C.int32_t) C.int32_t {
	ss := workbooks[h]
	sheet_name := C.GoString(sheet)
	sh, _ := ss.GetSheet(sheet_name)
	for i := 1; i < int(count); i++ {
		sh.AddRow()
	}
	return count
}

//export ss_insert_rows
func ss_insert_rows(h C.Handle, sheetName *C.char, rowNum, rows C.int32_t) C.ss_status {
	sheet, err := get_sheet(h, sheetName)
	if err != nil {
		return C.ss_worksheet_error
	}
	sheet.InsertRows(int(rowNum), uint32(rows))
	return C.ss_ok
}

//export ss_copy_rows
func ss_copy_rows(h C.Handle, sheetName *C.char, source, dest, rows C.int32_t) C.ss_status {
	sheet, err := get_sheet(h, sheetName)
	if err != nil {
		return C.ss_worksheet_error
	}
	sheet.CopyRows(uint32(source), uint32(dest), int(rows))
	return C.ss_ok
}

//export ss_auto_height
func ss_auto_height(h C.Handle, sheetName *C.char, row C.int32_t) C.ss_status {
	sheet, err := get_sheet(h, sheetName)
	if err != nil {
		return C.ss_worksheet_error
	}
	sheet.Row(uint32(row)).SetHeightAuto()
	return C.ss_ok
}

//export ss_add_cell
func ss_add_cell(h C.Handle, sheet *C.char, row C.uint32_t) *C.char {
	ss := workbooks[h]
	sheet_name := C.GoString(sheet)
	sh, _ := ss.GetSheet(sheet_name)
	r := sh.Row(uint32(row))
	cell := r.AddCell()
	return C.CString(cell.Reference())
}

//export ss_close
func ss_close(h C.Handle) {
	ss := workbooks[h]
	ss.Close()
	delete(workbooks, h)
}

//export ss_save
func ss_save(ws C.Handle, filepath *C.char) C.ss_status {
	wb := workbooks[ws]
	defer wb.Close()
	defer delete(workbooks, ws)
	err := wb.SaveToFile(C.GoString(filepath))
	if err != nil {
		return C.ss_save_failed
	}
	return C.ss_ok
}

//export ss_save_pdf
func ss_save_pdf(ws C.Handle, sheet, dest *C.char) C.ss_status {
	wb := workbooks[ws]
	defer wb.Close()
	defer delete(workbooks, ws)
	sheet_name := C.GoString(sheet)
	sh, err := wb.GetSheet(sheet_name)
	if err != nil {
		return C.ss_worksheet_error
	}
	pdf := convert2.ConvertToPdf(&sh)
	err = pdf.WriteToFile(C.GoString(dest))
	if err != nil {
		return C.ss_save_failed
	}
	return C.ss_ok
}

//export ss_check_sheet
func ss_check_sheet(h C.Handle, sheet *C.char) C.ss_status {
	_, err := get_sheet(h, sheet)
	if err != nil {
		return C.ss_worksheet_error
	}
	return C.ss_ok
}

func get_sheet(h C.Handle, sheet *C.char) (spreadsheet.Sheet, error) {
	ss := workbooks[h]
	sheet_name := C.GoString(sheet)
	return ss.GetSheet(sheet_name)
}

func get_cell(h C.Handle, sheet *C.char, cell *C.char) (spreadsheet.Cell, error) {
	sh, err := get_sheet(h, sheet)
	var c spreadsheet.Cell
	if err == nil {
		c = sh.Cell(C.GoString(cell))
	} else {
		setError(135)
	}
	return c, err
}

//export ss_set_cell_string
func ss_set_cell_string(h C.Handle, sheet *C.char, cell, value *C.char) C.int32_t {
	c, err := get_cell(h, sheet, cell)
	if err != nil {
		return 1
	}
	c.SetString(C.GoString(value))
	return 0
}

//export ss_set_cell_bool
func ss_set_cell_bool(h C.Handle, sheet *C.char, cell *C.char, value C.uint8_t) C.int32_t {
	c, err := get_cell(h, sheet, cell)
	if err != nil {
		return 1
	}
	c.SetBool(value == 1)
	return 0
}

//export ss_set_cell_date
func ss_set_cell_date(h C.Handle, sheet *C.char, cell *C.char, value C.double) C.int32_t {
	c, err := get_cell(h, sheet, cell)
	if err != nil {
		return 1
	}
	d := time.UnixMilli(int64(value))
	c.SetDate(d)
	return 0
}

//export ss_set_cell_date_with_style
func ss_set_cell_date_with_style(h C.Handle, sheet *C.char, cell *C.char, value C.double) C.int32_t {
	c, err := get_cell(h, sheet, cell)
	if err != nil {
		return 1
	}
	d := time.UnixMilli(int64(value))
	c.SetDateWithStyle(d)
	return 0
}

//export ss_set_cell_formula_array
func ss_set_cell_formula_array(h C.Handle, sheet *C.char, cell *C.char, value *C.char) C.int32_t {
	c, err := get_cell(h, sheet, cell)
	if err != nil {
		return 1
	}
	c.SetFormulaArray(C.GoString(value))
	return 0
}

//export ss_set_cell_formula_raw
func ss_set_cell_formula_raw(h C.Handle, sheet *C.char, cell *C.char, value *C.char) C.int32_t {
	c, err := get_cell(h, sheet, cell)
	if err != nil {
		return 1
	}
	c.SetFormulaRaw(C.GoString(value))
	return 0
}

//export ss_set_cell_formula_shared
func ss_set_cell_formula_shared(h C.Handle, sheet *C.char, cell *C.char, value *C.char, rows, cols C.uint32_t) C.int32_t {
	c, err := get_cell(h, sheet, cell)
	if err != nil {
		return 1
	}
	c.SetFormulaShared(C.GoString(value), uint32(rows), uint32(cols))
	return 0
}

//export ss_set_cell_number
func ss_set_cell_number(h C.Handle, sheet *C.char, cell *C.char, value C.double) C.int32_t {
	c, err := get_cell(h, sheet, cell)
	if err != nil {
		return 1
	}
	c.SetNumber(float64(value))
	return 0
}

// func (c Cell) SetHyperlink(hl common.Hyperlink)
// func (c Cell) SetRichTextString() RichText

type CellValue struct {
	V     string
	TAttr byte
}

//export ss_cell_get_value
func ss_cell_get_value(h C.Handle, sheet *C.char, cell *C.char) C.cellValue {
	c, err := get_cell(h, sheet, cell)
	if err != nil {
		return C.cellValue{}
	}

	raw, err := c.GetRawValue()
	if err != nil {
		setError(134)
		return C.cellValue{}
	}
	t := c.X().TAttr
	return C.cellValue{
		v: C.CString(raw),
		t: C.uint8_t(t),
		s: C.CString(c.GetFormat()),
	}
}

//export ss_cell_get_as_string
func ss_cell_get_as_string(h C.Handle, sheet *C.char, cell *C.char) *C.char {
	c, err := get_cell(h, sheet, cell)
	if err != nil {
		return C.CString("")
	}
	return C.CString(c.GetString())
}

//export ss_sheet_get_rows_as_strings
func ss_sheet_get_rows_as_strings(h C.Handle, sheet *C.char, row C.uint32_t) {

}

//export ss_cell_get_as_number
func ss_cell_get_as_number(h C.Handle, sheet *C.char, cell *C.char) C.double {
	c, _ := get_cell(h, sheet, cell)
	n, err := c.GetValueAsNumber()
	if err != nil {
		setError(134)
		return 0
	}
	return C.double(n)
}

//export ss_cell_get_bool
func ss_cell_get_bool(h C.Handle, sheet *C.char, cell *C.char) C.uint8_t {
	c, _ := get_cell(h, sheet, cell)
	v, err := c.GetValueAsBool()
	if err != nil {
		setError(134)
		return 0
	}
	if v {
		return 1
	}
	return 0
}

//export ss_cell_get_date
func ss_cell_get_date(h C.Handle, sheet *C.char, cell *C.char) C.int64_t {
	c, _ := get_cell(h, sheet, cell)
	v, err := c.GetValueAsTime2()

	if err != nil {
		s := c.GetString()
		fmt.Println("err", err, s)
		setError(134)
		return 0
	}
	return C.int64_t(v.UnixMilli())
}

//export ss_recalculate_formulas
func ss_recalculate_formulas(h C.Handle, sheet *C.char) {
	sh, err := get_sheet(h, sheet)
	if err == nil {
		sh.RecalculateFormulas()
	}
}

//export ss_last_column_index
func ss_last_column_index(h C.Handle, sheet *C.char) C.int32_t {
	sh, err := get_sheet(h, sheet)
	if err != nil {
		return 0
	}
	return C.int32_t(sh.MaxColumnIdx())
}

//export ss_last_row_index
func ss_last_row_index(h C.Handle, sheet *C.char) C.int32_t {
	sh, err := get_sheet(h, sheet)
	if err != nil {
		return 0
	}
	return C.int32_t(len(sh.Rows()))
}

//export ss_get_sheet_name
func ss_get_sheet_name(h C.Handle, sheet C.int32_t) *C.char {
	wb := workbooks[h]
	sheets := wb.Sheets()
	idx := int(sheet)
	if idx < 0 || idx >= len(sheets) {
		return nil
	}
	sh := sheets[idx]
	return C.CString(sh.Name())
}

var workbooks = make(map[C.Handle]*spreadsheet.Workbook)
