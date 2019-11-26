package main

import "fmt"
import "log"
import "os"
import "strings"
import "github.com/skratchdot/open-golang/open"
import "golang.org/x/sys/windows/registry"

func main() {
	fmt.Printf("Edition en ligne Silverpeas.\n")
	
	customURL := os.Args[1]
	
	log.Printf(customURL)
	
	customURL2 := strings.Replace(customURL, "spwebdav://", "http://", 1)
	webDavURL := strings.Replace(customURL2, "spwebdavs://", "https://", 1)
	
	openWithWindows(webDavURL)
}

func openWithWindows(url string) {
	log.Print("This is Windows !")

	parts := strings.Split(url, "/")
	fileName := parts[len(parts)-1]
	extension := strings.ToLower(substr(fileName,strings.LastIndex(fileName, "."),len(fileName)-strings.LastIndex(fileName, ".")))
//	log.Print("Extension = "+extension)

	if isMSProject(extension) {
		if isMSOfficeInstalled() {
			openWithMSOffice(url)
		}	else {
			err2 := open.Run(url)
			if err2 != nil {
				log.Fatal(err2)
			}			
		} 
	}	else {
		if (isOpenOffice(extension)) {
			openWithOpenOffice(url)
		} else { 
			if isMSOfficeInstalled() {
				openWithMSOffice(url)
			}	else { 
				err2 := open.Run(url)
				if err2 != nil {
					log.Fatal(err2)
				}			
			}
		}
	}
}

func isMSOfficeInstalled() bool {
	k, err := registry.OpenKey(registry.CLASSES_ROOT, "Word.Application", registry.QUERY_VALUE)
	if err == registry.ErrNotExist {
		log.Print(err)	
		return false
	} else {
		return true
	}
	defer k.Close()
	return false
}

func openWithOpenOffice(url string) {
	err2 := open.RunWith(url, "soffice.exe")
	if err2 != nil {
		log.Fatal(err2)
	}
}

func openWithMSOffice(url string) {
	err2 := open.RunWith(url, getMSApplication(url))
	if err2 != nil {
		log.Fatal(err2)
	}
}

func getMSApplication(url string) string {
	parts := strings.Split(url, "/")
	fileName := parts[len(parts)-1]
	extension := strings.ToLower(substr(fileName,strings.LastIndex(fileName, "."),len(fileName)-strings.LastIndex(fileName, ".")))
	
	if isPowerPoint(extension) {
		log.Print("Extension = "+extension+" -> powerpnt.exe")
		return "powerpnt.exe"
	} else if isExcel(extension) {
		log.Print("Extension = "+extension+" -> excel.exe")
		return "excel.exe"
	} else if isMSProject(extension) {
		log.Print("Extension = "+extension+" -> winproj.exe")
		return "winproj.exe"
	} else {
		log.Print("Extension = "+extension+" -> winword.exe")
		return "winword.exe"
	}
}
func substr(input string, start int, length int) string {
    asRunes := []rune(input)

    if start >= len(asRunes) {
        return ""
    }

    if start+length > len(asRunes) {
        length = len(asRunes) - start
    }
    return string(asRunes[start : start+length])
}

func isPowerPoint(extension string) bool {
	if strings.HasSuffix(extension, ".ppt") || strings.HasSuffix(extension, ".pot")	{
		return true
		}
		return false
}	

func isExcel(extension string) bool {
	if strings.HasPrefix(extension, ".xls") || strings.HasPrefix(extension, ".xlt") || strings.HasPrefix(extension, ".xlam") {
			return true
			}
		return false
}	

func isMSProject(extension string) bool {
	if strings.HasPrefix(extension, ".mpp") || strings.HasPrefix(extension, ".mpt") {
			return true
			}
	return false
}	

func isOpenOffice(extension string) bool {
	if strings.HasPrefix(extension, ".od") || strings.HasPrefix(extension, ".ot") {
			return true
			}
	return false
}	
