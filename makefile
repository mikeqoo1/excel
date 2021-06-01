#主程式
Main=main.go

#執行檔
Bin_Out= a.out 

main:
	go install
	go build -o ${Bin_Out} ${Main}
