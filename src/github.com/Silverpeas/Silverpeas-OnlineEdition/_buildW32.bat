set GOOS=windows
set GOARCH=386
go build -i -o onlineEditing32.exe onlineEditing_common.go onlineEditing_windows.go