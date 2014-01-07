goxls
=====

golang read xls(biff8)


func main() {
	//fmt.Println(BLANK_CELL)
	//readcell("test.xls", "bPC", 1, 1)
	for i := 0; i <= 63; i++ {
		readcell("test.xls", "AIO", uint16(i))
	}

}
