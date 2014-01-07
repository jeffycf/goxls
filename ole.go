package main

import (
	"encoding/binary"
	"fmt"
	"os"
)

const MSATSECT = 0xFFFFFFFC
const FATSECT = 0xFFFFFFFD
const ENDOFCHAIN = 0xFFFFFFFE
const FREESECT = 0xFFFFFFFF

func Ole_open(filename string) (xlsworkbook XlsWorkBook) {
	var oleh OLE2Header
	var ole OLE2
	buf := make([]byte, 512)
	file, err := os.Open(filename)
	if err != nil {
		fmt.Println(err)
	}
	ole.file = *file

	_, err = file.ReadAt(buf, 0)
	if err != nil {
		fmt.Println(err)
	}
	oleh.Setvalue(buf[0:512])
	if oleh.id != 0xD0CF11E0A1B11AE1 {
		fmt.Println("Not an excel file")
		os.Exit(1)
	}

	//fmt.Printf("%X\n", oleh)
	ole.lsector = 2 << (oleh.lsectorB - 1)
	ole.lssector = 2 << (oleh.lssectorB - 1)
	ole.cfat = oleh.cfat
	ole.dirstart = oleh.dirstart
	ole.sectorcutoff = oleh.sectorcutoff
	ole.sfatstart = oleh.sfatstart
	ole.csfat = oleh.csfat
	ole.difstart = oleh.difstart
	ole.cdif = oleh.cdif
	ole.files.count = 0

	//fmt.Printf("%X\n", oleh.MAST[0])
	read_MSAT(&ole, oleh)
	//fmt.Printf("%X\n", ole.SSecid)
	//fmt.Println(len((*ole.SSecid)))
	olest := ole2_sopen(&ole)
	allpss := ole2_dir(olest.buf)
	for _, val := range allpss {
		//fmt.Println(len(val.name))
		if changename(val.name) == "Workbook" {
			//fmt.Println(val.name)
			//workbook := ole2_stream(olest, val.sstart)
			//fmt.Printf("%X\n", workbook)
			xlsworkbook.olestr = ole2_stream(olest, val.sstart)
		}
	}
	return

}

func (oleh *OLE2Header) Setvalue(buff []byte) {

	oleh.id = binary.BigEndian.Uint64(buff[0:8])
	oleh.clid[0] = binary.BigEndian.Uint64(buff[8:16])
	oleh.clid[1] = binary.BigEndian.Uint64(buff[16:24])
	oleh.verminor = binary.LittleEndian.Uint16(buff[24:26])
	oleh.verdll = binary.LittleEndian.Uint16(buff[26:28])
	oleh.byteorder = binary.BigEndian.Uint16(buff[28:30])
	oleh.lsectorB = binary.LittleEndian.Uint16(buff[30:32])
	oleh.lssectorB = binary.LittleEndian.Uint16(buff[32:34])
	oleh.reserved1 = binary.LittleEndian.Uint16(buff[34:36])
	oleh.reserved2 = binary.LittleEndian.Uint32(buff[36:40])
	oleh.reserved3 = binary.LittleEndian.Uint32(buff[40:44])
	oleh.cfat = binary.LittleEndian.Uint32(buff[44:48])
	oleh.dirstart = binary.LittleEndian.Uint32(buff[48:52])
	oleh.reserved4 = binary.LittleEndian.Uint32(buff[52:56])
	oleh.sectorcutoff = binary.LittleEndian.Uint32(buff[56:60])
	oleh.sfatstart = binary.LittleEndian.Uint32(buff[60:64])
	oleh.csfat = binary.LittleEndian.Uint32(buff[64:68])
	oleh.difstart = binary.LittleEndian.Uint32(buff[68:72])
	oleh.cdif = binary.LittleEndian.Uint32(buff[72:76])

	for i := 0; i < 109; i = i + 1 {
		oleh.MAST[i] = binary.LittleEndian.Uint32(buff[76+i*4 : 76+i*4+4])
	}

}

func read_MSAT(ole2 *OLE2, oleh OLE2Header) {
	sectorNum := 109
	if ole2.cfat < 109 {
		sectorNum = int(ole2.cfat)
	}
	//read first 109 sectors of MSAT from header
	//index := ole2.cfat * uint32(ole2.lsector)
	//fmt.Println(sectorNum)
	var buff []uint32

	for i := 0; i < sectorNum; i++ {
		//fmt.Printf("%X\n", oleh.MAST[i])
		//fmt.Printf("%X\n", ole2.Secid)
		buff = append(buff, sector_read(ole2, oleh.MAST[i])...)
	}
	//fmt.Printf("%X\n", buff)
	ole2.Secid = &buff
	//Add additionnal sectors of the MSAT
	sid := ole2.difstart
	for sid != ENDOFCHAIN && sid != FREESECT {
		fmt.Printf("Add additionnal sectors of the MSAT:%X\n", sid)
	}

	//read in short table
	var buff2 []uint32
	if ole2.sfatstart != ENDOFCHAIN {
		//fmt.Println("read in short table")
		sec := ole2.sfatstart
		for k := 0; k < int(ole2.csfat); k++ {
			//fmt.Printf("%X\n", sec)
			tmp := (*ole2.Secid)[sec]
			//fmt.Printf("%X\n", tmp)
			buf := make([]byte, 512)
			pos := int64(sec*uint32(ole2.lsector) + uint32(512))
			//sec = (*ole2.Secid)[sec]
			_, err := ole2.file.ReadAt(buf, pos)
			if err != nil {
				fmt.Println(err)
			}
			for i := 0; i < 512; i = i + 4 {
				buff2 = append(buff2, binary.LittleEndian.Uint32(buf[i:i+4]))
			}
			sec = tmp
			//sec = (*ole2.Secid)[sec]
			//fmt.Printf("%X\n", k)
			//fmt.Printf("%X\n", buff2)
		}
	}
	ole2.SSecid = &buff2
	//fmt.Printf("%X\n", buff2)

}

func sector_read(ole2 *OLE2, sid uint32) (b []uint32) {
	buf := make([]byte, 512)
	_, err := ole2.file.ReadAt(buf, int64(sid)*512+512)
	if err != nil {
		fmt.Println(err)
	}
	for i := 0; i < 512; i = i + 4 {
		//fmt.Printf("%X\n", binary.LittleEndian.Uint32(buf[i:i+4]))
		b = append(b, binary.LittleEndian.Uint32(buf[i:i+4]))
	}
	//fmt.Printf("%X\n", buf)
	return
}

func ole2_sopen(ole *OLE2) (olest OLE2Stream) {
	//fmt.Println(ole.dirstart)
	//olest := OLE2Stream{}
	olest.ole = ole
	olest.size = -1
	olest.fatpos = uint32(ole.dirstart)
	olest.start = ole.dirstart
	olest.pos = 0
	olest.eof = 0
	olest.cfat = -1
	olest.bufsize = ole.lsector
	olest.buf = ole2_bufread(olest)
	//ole2_dir(olest.buf)
	//fmt.Printf("%X\n", olest.buf)
	return
}

func ole2_bufread(olest OLE2Stream) (buff []byte) {
	buf := make([]byte, 512)
	//buff := make([]byte, 512*4)
	for olest.fatpos != ENDOFCHAIN {
		_, err := olest.ole.file.ReadAt(buf, int64(olest.fatpos*uint32(512)+uint32(512)))
		if err != nil {
			fmt.Println(err)
		}
		//fmt.Printf("%X\n%X\n", olest.fatpos, (*olest.ole.Secid)[olest.fatpos])
		olest.fatpos = (*olest.ole.Secid)[olest.fatpos]
		buff = append(buff, buf...)
	}
	//fmt.Printf("%X\n", buff)
	//b = &buff
	return
}

func ole2_dir(buff []byte) (allpss []PSS) {
	pss := &PSS{}
	//allpss := []PSS{}
	for i := 0; i < len(buff); i = i + 128 {
		//fmt.Printf("%X\n", buff[i:i+128])

		pss.parsePSS(buff[i : i+128])
		allpss = append(allpss, *pss)
		//fmt.Println(string(pss.name))
		//fmt.Printf("%X\t%X\n", pss.sstart, pss.size)
	}
	//fmt.Printf("%X\n", allpss, len(allpss))
	return
}

func (p *PSS) parsePSS(buff []byte) {
	p.name = buff[0:64]
	p.bsize = buff[64:66]
	p.ttype = buff[66:67]
	p.flag = buff[67:68]
	p.left = buff[68:72]
	p.right = buff[72:76]
	p.child = buff[76:80]
	p.guid = buff[80:96]
	p.userflags = buff[96:100]
	p.time = buff[100:116]
	p.sstart = buff[116:120]
	p.size = buff[120:124]
	p.proptype = buff[124:128]
}

func changename(name []byte) (names string) {
	for _, val := range name {
		if val == 0 {
			continue
		}
		names = names + string(val)
	}
	return
}

func ole2_stream(olest OLE2Stream, start []byte) (buff []byte) {
	buf := make([]byte, 512)
	ss := binary.LittleEndian.Uint32(start)
	for ss != ENDOFCHAIN {
		_, err := olest.ole.file.ReadAt(buf, int64(ss*uint32(512)+uint32(512)))
		if err != nil {
			fmt.Println(err)
		}
		//fmt.Printf("%X\t %X\n", ss, (*olest.ole.Secid)[ss])
		ss = (*olest.ole.Secid)[ss]
		buff = append(buff, buf...)
	}
	//fmt.Printf("%X\n", buff)
	//b = &buff
	return
}
