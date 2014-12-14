package main
import (
		"fmt"
		//"os"
		"net/smtp"
		"github.com/tealeg/xlsx"
		"strings"
		"strconv"
       )

func main() {
		out := make([] string, 100)
		OutFileName := "out.xlsx"
		OutFile, OutError := xlsx.OpenFile(OutFileName)
		if OutError != nil {
		       //...
		}
		m := 0
		for _, OutSheet := range OutFile.Sheets {
			for _, OutRow := range OutSheet.Rows{
				for _, OutCell := range OutRow.Cells {
					out[m] = OutCell.String()
					m ++
				}
			}
		}
		// fmt.Println(out)
		// os.Exit(1)


		excelFileName := "foo.xlsx"
		xlFile, error := xlsx.OpenFile(excelFileName)
		if error != nil {
		       //...
		}
		all := 0
		for _, sheet := range xlFile.Sheets {
		       for _, row := range sheet.Rows {
				e := row.Cells
				user := e[0].String()
				password := e[1].String()
				host := e[2].String()	
				hp := strings.Split(host, ":")
				number,err := strconv.Atoi(e[3].String())
				if err != nil {
				       //...
				}
				// 发送限定数目的邮件
				for i:=0;i<=number;i++ {
					if len(out[all])>0 {
						to := out[all]
						
						subject := "测试3"
						body := `
							<html>
								<body>
									<h3>
								    "这是一封测试自动发送的邮件-26度C"
									</h3>
								</body>
							</html>
							`	
						auth := smtp.PlainAuth("", user, password, hp[0])
										
						var content_type string
						mailtype := "html"
						if mailtype == "html" {
							content_type = "Content-Type: text/"+ mailtype + "; charset=UTF-8"
						}else{
							content_type = "Content-Type: text/plain" + "; charset=UTF-8"
						}
						msg := []byte("To: " + to + "\r\nFrom: " + user + "<"+ user +">\r\nSubject: " + subject + "\r\n" + content_type + "\r\n\r\n" + body)
						send_to := strings.Split(to, ";")
						err := smtp.SendMail(host, auth, user, send_to, msg)
						fmt.Println(err)
					}
					all ++
				}
		       }
		}
	}
