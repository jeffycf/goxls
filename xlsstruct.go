package main

import (
	"encoding/binary"
	"os"
)

const PS_EMPTY = 00
const PS_USER_STORAGE = 01
const PS_USER_STREAM = 02
const PS_USER_ROOT = 05
const BLACK = 1

type BOF struct {
	id   uint16
	size uint16
}

type BIFF struct {
	ver     uint16
	btype   uint16
	id_make uint16
	year    uint16
	flags   uint32
	min_ver uint32
	buf     [100]byte
}

type WIND1 struct {
	xwn                             uint16
	ywn, dxwn, dywn, grbit, itabcur uint16
	itabfirst, ctabsel, wtabratio   uint16
}
type BOUNDSHEET struct {
	filepos uint32
	ttype   []byte
	visible []byte
	name    string
}

func (bound *BOUNDSHEET) setboundsheed(buf []byte) {
	bound.filepos = binary.LittleEndian.Uint32(buf[0:4])
	bound.ttype = buf[4:5]
	bound.visible = buf[5:6]
	//fmt.Println(len(buf))
	bound.name = string(buf[7:])

}

type ROW struct {
	index, fcell, lcell, height  uint16
	notused, notused2, flags, xf uint16
}

func (r *ROW) setRow(buf []byte) {
	r.index = binary.LittleEndian.Uint16(buf[0:2])
	r.fcell = binary.LittleEndian.Uint16(buf[2:4])
	r.lcell = binary.LittleEndian.Uint16(buf[4:6])
	r.height = binary.LittleEndian.Uint16(buf[6:8])
	r.flags = binary.LittleEndian.Uint16(buf[12:14])
	r.xf = binary.LittleEndian.Uint16(buf[14:16])
}

type COL struct {
	row, col, xf uint16
}

//func (c *COL)setCOL(buf []byte){
//	c.row=
//	c.col=
//	c.xf=
//}

type FORMULA struct {
	row, col, xf uint16
	resid        byte
	resdata      [5]byte
	res, flags   uint16
	chn          [4]byte
	llen         uint16
	value        [1]byte
}
type PK struct {
	row, col, xf uint16
	value        [1]byte
}

type BLANK struct {
	row, col, xf uint16
}
type LABEL struct {
	row, col, xf uint16
	value        uint32
}

func (label *LABELSST) setLabel(buf []byte) {
	label.row = binary.LittleEndian.Uint16(buf[0:2])
	label.col = binary.LittleEndian.Uint16(buf[2:4])
	label.xf = binary.LittleEndian.Uint16(buf[4:6])
	label.value = binary.LittleEndian.Uint32(buf[6:10])
}

type LABELSST LABEL

type SST struct {
	num, numofstr uint32
	buff          []byte
}

func (s *SST) setSst(buf []byte) {
	s.num = binary.LittleEndian.Uint32(buf[0:4])
	s.numofstr = binary.LittleEndian.Uint32(buf[4:8])
	s.buff = buf[8:]
}

type SST_DATA struct {
	str []string
}

func (sstdata *SST_DATA) setsstdata(sst SST) {
	var ss []string
	var s string
	//fmt.Println(len(sst.buff))
	for i := 0; i < len(sst.buff); {
		//fmt.Printf("%X\n", sst.buff[i:i+20])
		var fHighByte, fExtSt, fRichSt byte
		fHighByte = sst.buff[i+2] & 0x01
		fExtSt = sst.buff[i+2] & 0x04
		fRichSt = sst.buff[i+2] & 0x08
		//fmt.Println(sst[i+2], fHighByte, fExtSt, fRichSt)
		len_s := binary.LittleEndian.Uint16(sst.buff[i : i+2])
		s = string(sst.buff[i+3 : i+3+int(len_s)])
		var op, crun, cchExtRst uint16
		if fExtSt == 4 && fRichSt == 8 {
			crun = binary.LittleEndian.Uint16(sst.buff[i+3 : i+5])
			cchExtRst = binary.LittleEndian.Uint16(sst.buff[i+5 : i+9])
			op = op + 2 + 4
		} else if fExtSt == 4 && fRichSt != 8 {
			cchExtRst = binary.LittleEndian.Uint16(sst.buff[i+3 : i+7])
			op = op + 4
		} else if fRichSt == 8 && fExtSt != 4 {
			crun = binary.LittleEndian.Uint16(sst.buff[i+3 : i+5])
			op = op + 2

		} else {
			cchExtRst = 0
			op = op
			crun = 0
		}
		//fmt.Println(i, op, crun, cchExtRst, fHighByte, fExtSt, fRichSt, len_s)
		if fHighByte == 0 {
			s = string(sst.buff[i+3+int(op) : i+3+int(len_s)+int(op)])
			//fmt.Println(s)
			i = i + int(len_s) + 3 + int(op) + int(crun)*4 + int(cchExtRst)
		} else if fHighByte == 1 {
			i = i + int(len_s)*2 + 3 + int(op) + int(crun)*4 + int(cchExtRst)
		}
		ss = append(ss, s)
	}
	sstdata.str = ss
	//fmt.Println(ss[len(ss)-1], len(ss))
}

type XF5 struct {
	font, format, ttype, align, color, fill, border, linestyle uint16
}

type XF8 struct {
	font, format, ttype              uint16
	align, rotation, ident, usedattr byte
	linestyle, linecolor             uint32
	groundcolor                      uint16
}
type BR_NUMBER struct {
	row, col, xf uint16
	value        [8]byte
}

type COLINFO struct {
	first, last, width, xf, flags, notused uint16
}

type MERGEDCELLS struct {
	rowf, rowl, colf, coll uint16
}

type FONT struct {
	height, flag, color, bold, escapement     uint16
	underline, family, charset, notused, name byte
}
type FORMAT struct {
	index uint16
	value [0]byte
}

type st_sheet_data struct {
	filepos           uint32
	visibility, ttype []byte
	name              string
}

type St_sheet struct {
	count int
	sheet []st_sheet_data
}
type St_font struct {
	count uint32
	font  *sf_font_data
}
type sf_font_data struct {
	height, flag, color, bold, escapement uint16
	underline, family, charset            byte
	name                                  *byte
}
type st_fromat_data struct {
	index uint16
	value *byte
}
type St_format struct {
	count  uint32
	format *st_fromat_data
}

type st_xf_data struct {
	font, format, ttype              uint16
	align, rotation, ident, usedattr byte
	linestyle, linecolor             uint32
	groundcolor                      uint16
}

type St_xf struct {
	count uint32
	xf    *st_xf_data
}

type St_sst struct {
	count, lastid, continued, lastln, lastrt, lastsz uint32
	sst_string                                       string
}
type st_cell_data struct {
	id, row, col, xf        uint16
	str                     string
	d                       float64
	l                       uint16
	width, colspan, rowspan uint16
	ishidden                byte
}
type St_cell struct {
	count uint16
	cell  *st_cell_data
}
type st_row_data struct {
	index, fcell, lcell, height, flags, xf uint16
	xfflags                                byte
	cells                                  St_cell
}

type St_row struct {
	//count uint32
	lastcol, lastrow uint16 //numcols-1,numrows-1
	row              []st_row_data
}
type St_colinfo struct {
	//count uint32
	col []st_colinfo_data
}
type st_colinfo_data struct {
	first, last, width, xf, flags uint16
}

type XlsWorkBook struct {
	//file File*
	olestr     []byte
	filepos    uint16
	is5ver     byte
	is1904     byte
	ttype      uint16
	codepage   uint16
	charset    string
	sheets     []st_sheet_data
	sst        SST_DATA
	xfs        St_xf
	fonts      St_font
	formats    St_format
	summary    string
	docsummary string
	worksheets []XlsWorkSheet
}
type Xls_summaryinfo struct {
	title, subject, author, keywords, comment       string
	lastAuthor, appname, category, manager, company string
}
type XlsWorkSheet struct {
	filepos     uint32
	defcolwidth uint16
	frow, lrow  uint32
	fcol, lcol  uint16
	rows        []St_row
	//workbook    *XlsWorkBook
	colinfo []St_colinfo
	lsst    []LABELSST
	name    string
}

//ole struct
type TIME_T struct {
	LowDate, HighDate uint32
}

type OLE2Header struct {
	id                                                            uint64 //D0CF11E0 A1B11AE1
	clid                                                          [2]uint64
	verminor, verdll, byteorder, lsectorB, lssectorB, reserved1   uint16
	reserved2, reserved3, cfat, dirstart, reserved4, sectorcutoff uint32
	sfatstart, csfat, difstart, cdif                              uint32
	MAST                                                          [109]uint32
}

type st_olefiles struct {
	count uint16
	file  *st_olefiles_data
}

type st_olefiles_data struct {
	name        string
	start, size uint32
}
type OLE2 struct {
	file                                    os.File
	lsector, lssector                       uint16
	cfat, dirstart, sectorcutoff, sfatstart uint32
	csfat, difstart, cdif                   uint32
	Secid                                   *[]uint32
	SSecid                                  *[]uint32
	SSAT                                    *byte
	files                                   st_olefiles
}

type OLE2Stream struct {
	ole             *OLE2
	start, fatpos   uint32
	pos, cfat, size int16
	buf             []byte
	bufsize         uint16
	eof             byte
	sfat            byte
}
type PSS struct {
	name                   []byte //64
	bsize                  []byte //2
	ttype                  []byte //1
	flag                   []byte //1
	left, right, child     []byte //4
	guid                   []byte //16
	userflags              []byte //4
	time                   []byte //8
	sstart, size, proptype []byte //4
}
