package main

import (
	"encoding/binary"
	"fmt"
)

const BLANK_CELL = 0x201

type SectionList struct {
	format [4]uint32
	offset uint32
}

type Header struct {
	Sig     uint16
	_empty  uint16
	Os      uint32
	format  [4]uint32
	count   uint32
	SecList [0]SectionList
}

type PropertyList struct {
	PropertyID    uint32
	SectionOffset uint32
}

type SectionHeader struct {
	Length        uint32
	NumProperties uint32
	Properties    [0]PropertyList
}

type Property struct {
	ProperID uint32
	Data     [0]uint32
}

func main() {
	//fmt.Println(BLANK_CELL)
	//readcell("test.xls", "bPC", 1, 1)
	for i := 0; i <= 63; i++ {
		readcell("test.xls", "AIO", uint16(i))
	}

}

func readcell(filename, sheet string, row uint16) {
	workbook := Ole_open(filename)
	workbook.Xls_parseWorkBook()
	workbook.xls_parseWorkSheet()

	for _, val := range workbook.worksheets {
		//fmt.Println(val.name)
		if val.name[1:4] == sheet {
			for _, val1 := range val.lsst {
				//fmt.Println(val1.row, val1.col)
				if val1.row == row {
					//fmt.Println(val1.row, val1.col)
					//fmt.Println("here3!")
					fmt.Printf("%s\t", workbook.sst.str[int(val1.value)])
				}

			}
			fmt.Println()
			break

		}
	}

}

func Xls_open(filename string) {
	//workbook := Ole_open(filename)
	//fmt.Println(string(workbook.olestr[0:100]))
	//Xls_parseWorkBook(workbook)
}

func (workbook *XlsWorkBook) Xls_parseWorkBook() {
	var bof1 BOF
	//var bof2 BOF
	buf := make([]byte, 512)
	var once int
	//fmt.Println(len(workbook.olestr))
	var bounds []BOUNDSHEET
	sst := &SST{}
	for i := 0; i < len(workbook.olestr); {
		bof1.id = binary.LittleEndian.Uint16(workbook.olestr[i : i+2])
		bof1.size = binary.LittleEndian.Uint16(workbook.olestr[i+2 : i+4])
		buf = workbook.olestr[i+4 : i+int(bof1.size)+4]
		switch bof1.id {
		case 0x000A: //EOF
			i = i + 2
			break
		case 0x0809: //biff5-8
			//fmt.Printf("%X\n", buf)

		case 0x00E1: //interfachder
			//fmt.Printf("%X\n", buf)

		//case 0x0042: //codepage
		case 0x003c: //CONTINUE
			if once == 0 {
				if buf[0] == 0 || buf[0] == 1 {
					sst.buff = append(sst.buff, buf[1:]...)
				} else {
					sst.buff = append(sst.buff, buf...)
				}

			}
		//case 0x003d: //WINDOWS1
		case 0x00fc: //SST
			//fmt.Printf("%X\n", bof1.size)
			sst.setSst(buf)
			once = 0
			//fmt.Printf("%X\n", sst)
		case 0x00ff: //extsst

		case 0x0085: //boundsheet
			//fmt.Printf("%X\t%X\n", bof1, buf)
			var bound BOUNDSHEET
			bound.setboundsheed(buf)
			bounds = append(bounds, bound)
			//fmt.Printf("%X\n", bound.filepos)
			workbook.xls_addSheet(bound)

		//case 0x00e0: //xf
		//case 0x0031: //font
		//case 0x041e: //format
		//case 0x0293: //style
		//case 0x0092: //palette
		//case 0x0022: //1904
		//case 0x00eb:
		case 0x01b6:
			once = 1
		default:
			//fmt.Println("default", i)
			//fmt.Printf("%X\n", workbook.olestr[i:i+2])
			//fmt.Printf("%X\n", buf)
			//i = i + 4 + int(bof1.size)
		}
		if bof1.id != 0x000A {
			i = i + 4 + int(bof1.size)
		}

		//bof2 = bof1
		//once = 1
	}
	//fmt.Println(workbook.sheets)
	//fmt.Println(string(sst.buff))
	sstdata := &SST_DATA{}
	sstdata.setsstdata(*sst)
	workbook.sst = *sstdata
	workbook.xls_parseWorkSheet()
}

func (xlsworkbook *XlsWorkBook) xls_addSheet(bs BOUNDSHEET) {
	//st_sheet := &St_sheet{}
	st_sheet_data := st_sheet_data{}
	st_sheet_data.filepos = bs.filepos
	st_sheet_data.visibility = bs.visible
	st_sheet_data.ttype = bs.ttype
	st_sheet_data.name = bs.name
	xlsworkbook.sheets = append(xlsworkbook.sheets, st_sheet_data)
}

func (workbook *XlsWorkBook) xls_parseWorkSheet() {
	buf := make([]byte, 512)
	var bof1 BOF
	for _, val := range workbook.sheets {
		worksheet := &XlsWorkSheet{}
		worksheet.name = val.name
		worksheet.filepos = val.filepos
		//fmt.Println(len(val.name), val.name[1:4])
		//if val.name[1:4] != "bPC" {
		//	continue
		//}
		//fmt.Printf("%X\n", worksheet.filepos)
		i := int(worksheet.filepos)
		for i < len(workbook.olestr) {
			bof1.id = binary.LittleEndian.Uint16(workbook.olestr[i : i+2])
			bof1.size = binary.LittleEndian.Uint16(workbook.olestr[i+2 : i+4])
			buf = workbook.olestr[i+4 : i+int(bof1.size)+4]
			switch bof1.id {
			case 0x000A: //EOF
				i = i + 2
				break
			//case 0x00e5: //mergedcessl
			case 0x0208: //row
				row := &ROW{}
				row.setRow(buf)

			//case 0x0055:
			//case 0x0225:
			case 0x00d7:
				//fmt.Printf("%X\n", buf)
			case 0x020b:
				//fmt.Printf("%X\n", buf)

			//case 0x00BD: //MULRK
			//case 0x00BE: //MULBLANK
			//case 0x0203: //NUMBER
			//case 0x027e: //RK
			case 0x00FD: //LABELSST
				//fmt.Printf("%X\n", buf)
				lsst := &LABELSST{}
				lsst.setLabel(buf)
				worksheet.lsst = append(worksheet.lsst, *lsst)
				//fmt.Printf("%X\n", lsst)
			//case 0x0201: //BLANK
			//case 0x0204: //LABEL
			//case 0x0006: //FORMULA
			//case 0x0207:
			//case 0x01b8: //hyperref
			//case 0x023e: //windows2
			case 0x0200:
				worksheet.frow = binary.LittleEndian.Uint32(buf[0:4])
				worksheet.lrow = binary.LittleEndian.Uint32(buf[4:8])
				worksheet.fcol = binary.LittleEndian.Uint16(buf[8:10])
				worksheet.lcol = binary.LittleEndian.Uint16(buf[10:12])
			default:
			}
			if bof1.id != 0x000a {
				i = i + 4 + int(bof1.size)
			}
		}
		//fmt.Println(worksheet)
		workbook.worksheets = append(workbook.worksheets, *worksheet)
	}
	//fmt.Println(workbook.worksheets)
}
