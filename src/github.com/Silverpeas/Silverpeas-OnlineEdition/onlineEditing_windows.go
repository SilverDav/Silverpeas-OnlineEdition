package main

import "fmt"
import "log"
import "os"
import "strings"
import "github.com/skratchdot/open-golang/open"
import "golang.org/x/sys/windows/registry"

func main() {
	fmt.Println("Edition en ligne Silverpeas.\n")
	
	customURL := os.Args[1]
	
	log.Println(customURL)
	
	customURL2 := strings.Replace(customURL, "spwebdav://", "http://", 1)
	webDavURL := strings.Replace(customURL2, "spwebdavs://", "https://", 1)
	
	openWithWindows(webDavURL)
//  var input string
//  fmt.Scanln(&input)

}

func openWithWindows(url string) {
	log.Printf("Windows...")

	parts := strings.Split(url, "/")
	fileName := parts[len(parts)-1]
	extension := substr(fileName,strings.LastIndex(fileName, "."),len(fileName)-strings.LastIndex(fileName, "."))

	if isMSProject(extension) {
		if isMSProjectInstalled() {
			log.Printf("Ouvre avec MSProject...")
			openWithMSSoftware(url)
		}
	}	else { 
		if isMSOfficeInstalled() {
			log.Printf("Ouvre avec MSOffice...")
			openWithMSSoftware(url)
		}	else { 
			log.Printf("Ouvre avec soffice...")
			openWithOpenOffice(url)
		}
	}
}

func isMSProjectInstalled() bool {
	k, err := registry.OpenKey(registry.CLASSES_ROOT, "MSProject.Application", registry.QUERY_VALUE)
	if err == registry.ErrNotExist {
		fmt.Print(err)	
		return false
	} else {
		return true
	}
	defer k.Close()
	return false
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

func openWithMSSoftware(url string) {
	err2 := open.RunWith(url, getMSApplication(url))
	if err2 != nil {
		log.Fatal(err2)
	}
}

func getMSApplication(url string) string {
	parts := strings.Split(url, "/")
	fileName := parts[len(parts)-1]
	extension := substr(fileName,strings.LastIndex(fileName, "."),len(fileName)-strings.LastIndex(fileName, "."))
	
	if isPowerPoint(extension) {
		log.Printf("Extension = "+extension+" -> powerpnt.exe")
		return "powerpnt.exe"
	} else if isExcel(extension) {
		log.Printf("Extension = "+extension+" -> excel.exe")
		return "excel.exe"
	} else if isMSProject(extension) {
		log.Printf("Extension = "+extension+" -> winproj.exe")
		return "winproj.exe"
	} else {
		log.Printf("Extension = "+extension+" -> winword.exe")
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
	if strings.HasSuffix(extension, "ppt") || strings.HasSuffix(extension, "pot") || strings.HasSuffix(extension, "pptx") || strings.HasSuffix(extension, "pptm") || strings.HasSuffix(extension, "potx") || strings.HasSuffix(extension, "potm")	{
		return true
		}
		return false
}	

func isExcel(extension string) bool {
	if strings.HasSuffix(extension, "xls") || strings.HasSuffix(extension, "xlsx") || strings.HasSuffix(extension, "xlt") || strings.HasSuffix(extension, "xlsm") || strings.HasSuffix(extension, "xltx") || strings.HasSuffix(extension, "xltm")	 || strings.HasSuffix(extension, "xlsb")  || strings.HasSuffix(extension, "xlam") {
			return true
			}
		return false
}	

func isMSProject(extension string) bool {
	if strings.HasSuffix(extension, "mpp") || strings.HasSuffix(extension, "mpt") {
			return true
			}
	return false
}	
