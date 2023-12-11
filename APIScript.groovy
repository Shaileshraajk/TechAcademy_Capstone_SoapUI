import org.apache.poi.xssf.usermodel.XSSFWorkbook

// Excel file
def excelFilePath = "D:/TestData/Capstone Project - API Data.xlsx"

// Load Excel workbook and sheet
def excelFile = new File(excelFilePath)
def workbook = new XSSFWorkbook(excelFile)
def sheet = workbook.getSheet("Sheet1")


for (int row = 1; row <= sheet.getLastRowNum(); row++) {
	
    def name = sheet.getRow(row).getCell(0).getStringCellValue()
    def salary = sheet.getRow(row).getCell(1).getStringCellValue()
    def age = sheet.getRow(row).getCell(2).getStringCellValue()


    def request = """
        {
            "name": "${name}",
            "salary": "${salary}",
            "age": "${age}"
        }
    """

    // POST request
    def response = testRunner.testCase.testSteps["REST Request - Post"].run(request)

    // Assert response
    assert response.getResponseStatusCode() == 200

    def jsonResponse = new groovy.json.JsonSlurper().parseText(response.getResponseContent())
    assert jsonResponse.status == "success"
    assert jsonResponse.data.name == name
    assert jsonResponse.data.salary == salary
    assert jsonResponse.data.age == age

}

workbook.close()
