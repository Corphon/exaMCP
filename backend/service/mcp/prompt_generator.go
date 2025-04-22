package mcp

import (
	"bytes"
	"fmt"
	"strings"
	"text/template"
	"time"
)

// DataType represents Excel column data types
type DataType string

const (
	TypeText     DataType = "Text"
	TypeNumber   DataType = "Number"
	TypeDate     DataType = "Date"
	TypeBoolean  DataType = "Boolean"
	TypeCurrency DataType = "Currency"
	TypeFormula  DataType = "Formula"
	TypeUnknown  DataType = "Unknown"
)

// DataRange represents an Excel data range structure
type DataRange struct {
	RangeAddress string            // e.g., "A1:D10"
	Headers      []string          // Column headers
	DataRows     int               // Number of data rows
	DataTypes    map[string]string // Data types for each column
	SampleData   [][]string        // Sample data (max 5 rows)
	Description  string            // Auto-detected description of data
	HasHeaders   bool              // Whether the range has headers
	SheetName    string            // Sheet name
	Relationships []Relationship   // Related ranges
}

// Relationship represents a relationship between data ranges
type Relationship struct {
	TargetRange string // Target range reference
	Type        string // Relationship type (OneToMany, ManyToOne, etc.)
	SourceField string // Source field/column
	TargetField string // Target field/column
}

// PromptConfig contains configuration options for prompt generation
type PromptConfig struct {
	Language            string            // Prompt language (default: "en")
	IncludeExamples     bool              // Include example VBA code
	UseAdvancedContext  bool              // Include advanced context like relationships
	MaxSampleRows       int               // Maximum number of sample data rows to include
	HighlightKeyColumns []string          // Column names to highlight as important
	TemplateVariables   map[string]string // Custom template variables
	OutputType          string            // Type of output (Generic, DataProcessing, Reporting, etc.)
	DetailLevel         string            // Level of detail (Basic, Intermediate, Advanced)
	IncludeModules      []string          // Standard modules to include
	TargetExcelVersion  string            // Target Excel version
}

// DefaultPromptConfig returns default configuration for prompt generation
func DefaultPromptConfig() PromptConfig {
	return PromptConfig{
		Language:            "en",
		IncludeExamples:     true,
		UseAdvancedContext:  false,
		MaxSampleRows:       3,
		HighlightKeyColumns: []string{},
		TemplateVariables:   map[string]string{},
		OutputType:          "Generic",
		DetailLevel:         "Intermediate",
		IncludeModules:      []string{},
		TargetExcelVersion:  "Excel 2016+",
	}
}

// GenerateMCPPrompt generates a basic MCP-formatted prompt for VBA code generation
func GenerateMCPPrompt(structure DataRange, userRequirement string, includeStandardModules bool) string {
	config := DefaultPromptConfig()
	config.IncludeModules = getModuleList(includeStandardModules)
	return GenerateExaMCPPromptWithConfig(structure, userRequirement, config)
}

// getModuleList returns the list of standard modules based on the includeStandardModules flag
func getModuleList(includeStandardModules bool) []string {
	if includeStandardModules {
		return []string{"SQLUtils", "DataTools", "UIHelpers"}
	}
	return []string{}
}

// GenerateExaMCPPromptWithConfig generates a customized MCP prompt with the provided configuration
func GenerateExaMCPPromptWithConfig(structure DataRange, userRequirement string, config PromptConfig) string {
	var prompt strings.Builder

	// Select template based on configuration
	templateContent := getPromptTemplate(config)

	// Prepare template data
	data := map[string]interface{}{
		"CurrentDateTime": time.Now().Format("2006-01-02 15:04:05"),
		"Structure":       structure,
		"UserRequirement": userRequirement,
		"Config":          config,
		"HeadersFormatted": formatHeaders(structure.Headers, structure.DataTypes),
		"SampleDataLimited": limitSampleData(structure.SampleData, config.MaxSampleRows),
		"RelationshipDescriptions": formatRelationships(structure.Relationships),
		"ModuleDescriptions": getModuleDescriptions(config.IncludeModules),
		"Examples": getExamples(config),
		"ColumnLetters": generateColumnLetters(len(structure.Headers)),
	}

	// Add custom template variables
	for key, value := range config.TemplateVariables {
		data[key] = value
	}

	// Parse and execute template
	tmpl, err := template.New("mcp_prompt").Parse(templateContent)
	if err != nil {
		return fallbackPrompt(structure, userRequirement, config, err)
	}

	var buf bytes.Buffer
	if err := tmpl.Execute(&buf, data); err != nil {
		return fallbackPrompt(structure, userRequirement, config, err)
	}

	prompt.WriteString(buf.String())
	return prompt.String()
}

// getPromptTemplate returns the appropriate template based on the configuration
func getPromptTemplate(config PromptConfig) string {
	// Use the appropriate template based on configuration
	if config.OutputType == "Reporting" {
		return reportingTemplate
	} else if config.OutputType == "DataProcessing" {
		return dataProcessingTemplate
	} else if config.OutputType == "UserInterface" {
		return userInterfaceTemplate
	} else if config.UseAdvancedContext {
		return advancedTemplate
	}
	
	// Default template
	return basicTemplate
}

// formatHeaders formats the headers with their data types for the prompt
func formatHeaders(headers []string, dataTypes map[string]string) string {
	var result strings.Builder
	
	for i, header := range headers {
		dataType := dataTypes[header]
		if dataType == "" {
			dataType = "Unknown"
		}
		
		result.WriteString(fmt.Sprintf("[header:%s] (Column %s, Type: %s)\n", 
			header, columnLetterFromIndex(i), dataType))
	}
	
	return result.String()
}

// limitSampleData limits the sample data to the specified number of rows
func limitSampleData(sampleData [][]string, maxRows int) [][]string {
	if len(sampleData) <= maxRows {
		return sampleData
	}
	return sampleData[:maxRows]
}

// formatRelationships formats relationship descriptions for the prompt
func formatRelationships(relationships []Relationship) string {
	if len(relationships) == 0 {
		return "No relationships detected."
	}
	
	var result strings.Builder
	for i, rel := range relationships {
		result.WriteString(fmt.Sprintf("%d. %s related to %s through %s → %s\n", 
			i+1, rel.Type, rel.TargetRange, rel.SourceField, rel.TargetField))
	}
	
	return result.String()
}

// getModuleDescriptions returns descriptions of the specified modules
func getModuleDescriptions(modules []string) string {
	if len(modules) == 0 {
		return ""
	}
	
	var result strings.Builder
	
	moduleDescriptions := map[string]string{
		"SQLUtils": `The SQLUtils module provides SQL-like query capabilities for Excel data:
- getSQL(StrSQL As String, rng As Range, Optional title As Boolean = True): Executes SQL queries against Excel data
- Example: Call getSQL("SELECT [Name], SUM([Sales]) FROM [Sheet1$] GROUP BY [Name]", Sheet2.Range("A1"), True)`,

		"DataTools": `The DataTools module provides data manipulation utilities:
- FindRow(searchRange As Range, searchValue As Variant) As Long: Finds a row by value
- CopyRangeToSheet(sourceRange As Range, targetSheet As Worksheet, targetCell As String): Copies data between sheets
- SortRange(dataRange As Range, sortColumn As Integer, ascending As Boolean): Sorts a data range`,

		"UIHelpers": `The UIHelpers module helps with creating user interfaces:
- CreateInputForm(title As String, fields As Variant) As Variant: Creates a simple input form and returns values
- ShowProgressBar(title As String, max As Long): Shows a progress bar dialog
- UpdateProgress(value As Long): Updates the progress bar
- CloseProgressBar(): Closes the progress bar`,
	}
	
	for _, module := range modules {
		if desc, ok := moduleDescriptions[module]; ok {
			result.WriteString(fmt.Sprintf("## %s Module\n%s\n\n", module, desc))
		}
	}
	
	return result.String()
}

// getExamples returns example VBA code relevant to the configuration
func getExamples(config PromptConfig) string {
	if !config.IncludeExamples {
		return ""
	}
	
	// Select appropriate examples based on configuration
	var examples []string
	
	if config.OutputType == "Reporting" {
		examples = append(examples, reportingExample)
	} else if config.OutputType == "DataProcessing" {
		examples = append(examples, dataProcessingExample)
	} else if config.OutputType == "UserInterface" {
		examples = append(examples, userInterfaceExample)
	} else {
		// Default example
		examples = append(examples, basicExample)
	}
	
	// If SQL module is included, add SQL example
	if contains(config.IncludeModules, "SQLUtils") {
		examples = append(examples, sqlExample)
	}
	
	// Combine examples
	var result strings.Builder
	for i, example := range examples {
		result.WriteString(fmt.Sprintf("### Example %d\n%s\n\n", i+1, example))
	}
	
	return result.String()
}

// generateColumnLetters generates Excel column letters (A, B, C, ..., AA, AB, etc.)
func generateColumnLetters(count int) []string {
	letters := make([]string, count)
	for i := 0; i < count; i++ {
		letters[i] = columnLetterFromIndex(i)
	}
	return letters
}

// columnLetterFromIndex converts a 0-based index to Excel column letter
func columnLetterFromIndex(index int) string {
	result := ""
	for index >= 0 {
		remainder := index % 26
		result = string(rune('A'+remainder)) + result
		index = index/26 - 1
	}
	return result
}

// fallbackPrompt generates a simple prompt when template processing fails
func fallbackPrompt(structure DataRange, userRequirement string, config PromptConfig, err error) string {
	var prompt strings.Builder
	
	prompt.WriteString("# TASK: Generate Excel VBA script based on user requirements\n\n")
	prompt.WriteString(fmt.Sprintf("ERROR: Template processing failed: %v. Using fallback prompt.\n\n", err))
	
	// Basic structure information
	prompt.WriteString("## EXCEL STRUCTURE\n")
	prompt.WriteString(fmt.Sprintf("- Sheet: %s\n", structure.SheetName))
	prompt.WriteString(fmt.Sprintf("- Range: %s\n", structure.RangeAddress))
	prompt.WriteString(fmt.Sprintf("- Headers: %s\n", strings.Join(structure.Headers, ", ")))
	prompt.WriteString(fmt.Sprintf("- Data Rows: %d\n\n", structure.DataRows))
	
	// User requirement
	prompt.WriteString("## USER REQUIREMENT\n")
	prompt.WriteString(userRequirement)
	prompt.WriteString("\n\n")
	
	// Output instructions
	prompt.WriteString("## OUTPUT INSTRUCTIONS\n")
	prompt.WriteString("Generate VBA code that fulfills the user requirement.\n")
	
	return prompt.String()
}

// contains checks if a string slice contains a specific value
func contains(slice []string, value string) bool {
	for _, item := range slice {
		if item == value {
			return true
		}
	}
	return false
}

// Template definitions follow

// basicTemplate is the standard MCP prompt template
const basicTemplate = `# TASK: Generate Excel VBA script based on user requirements

## SYSTEM INFORMATION
- Framework: exaMCP (Excel Automation with Model Context Protocol)
- Timestamp: {{.CurrentDateTime}}
- Target Excel Version: {{.Config.TargetExcelVersion}}

## EXCEL STRUCTURE
- Sheet: {{.Structure.SheetName}}
- Range: {{.Structure.RangeAddress}}
- Total Rows: {{.Structure.DataRows}}
- Has Headers: {{.Structure.HasHeaders}}
- Description: {{.Structure.Description}}

## HEADERS
{{.HeadersFormatted}}

## SAMPLE DATA
{{range $index, $row := .SampleDataLimited}}Row {{add $index 1}}: {{join $row ", "}}
{{end}}

{{if .ModuleDescriptions}}
## STANDARD MODULES AVAILABLE
{{.ModuleDescriptions}}
{{end}}

{{if .Config.IncludeExamples}}
## EXAMPLES
{{.Examples}}
{{end}}

## USER REQUIREMENT
{{.UserRequirement}}

## OUTPUT INSTRUCTIONS
1. Analyze the Excel structure and user requirement carefully
2. Generate complete, working VBA code that fulfills the requirement
3. Include proper error handling
4. Use descriptive variable names and add comments to explain the logic
5. If standard modules are available, utilize them when appropriate
6. Return only the VBA code, without additional explanations
`

// advancedTemplate includes more contextual information for complex scenarios
const advancedTemplate = `# TASK: Generate Excel VBA script based on user requirements

## SYSTEM INFORMATION
- Framework: exaMCP (Excel Automation with Model Context Protocol)
- Timestamp: {{.CurrentDateTime}}
- Target Excel Version: {{.Config.TargetExcelVersion}}
- Detail Level: {{.Config.DetailLevel}}

## EXCEL STRUCTURE
- Sheet: {{.Structure.SheetName}}
- Range: {{.Structure.RangeAddress}}
- Total Rows: {{.Structure.DataRows}}
- Has Headers: {{.Structure.HasHeaders}}
- Description: {{.Structure.Description}}

## HEADERS
{{.HeadersFormatted}}

## KEY COLUMNS
{{range .Config.HighlightKeyColumns}}
- {{.}}: Critical for business logic
{{end}}

## SAMPLE DATA
{{range $index, $row := .SampleDataLimited}}Row {{add $index 1}}: {{join $row ", "}}
{{end}}

## DATA RELATIONSHIPS
{{.RelationshipDescriptions}}

{{if .ModuleDescriptions}}
## STANDARD MODULES AVAILABLE
{{.ModuleDescriptions}}
{{end}}

{{if .Config.IncludeExamples}}
## EXAMPLES
{{.Examples}}
{{end}}

## USER REQUIREMENT
{{.UserRequirement}}

## OUTPUT INSTRUCTIONS
1. Analyze the Excel structure, relationships, and user requirement carefully
2. Generate complete, working VBA code that fulfills the requirement
3. Include comprehensive error handling and validation
4. Use descriptive variable names and add detailed comments
5. Utilize standard modules when appropriate
6. Consider performance optimization for large datasets
7. Return only the VBA code, without additional explanations
`

// reportingTemplate is specialized for reporting tasks
const reportingTemplate = `# TASK: Generate Excel VBA reporting script

## SYSTEM INFORMATION
- Framework: exaMCP (Excel Automation with Model Context Protocol)
- Timestamp: {{.CurrentDateTime}}
- Target Excel Version: {{.Config.TargetExcelVersion}}
- Output Type: Reporting

## SOURCE DATA
- Sheet: {{.Structure.SheetName}}
- Range: {{.Structure.RangeAddress}}
- Total Rows: {{.Structure.DataRows}}
- Has Headers: {{.Structure.HasHeaders}}
- Description: {{.Structure.Description}}

## COLUMNS FOR REPORTING
{{.HeadersFormatted}}

## SAMPLE DATA
{{range $index, $row := .SampleDataLimited}}Row {{add $index 1}}: {{join $row ", "}}
{{end}}

{{if .ModuleDescriptions}}
## STANDARD MODULES AVAILABLE
{{.ModuleDescriptions}}
{{end}}

{{if .Config.IncludeExamples}}
## EXAMPLES
{{.Examples}}
{{end}}

## REPORTING REQUIREMENTS
{{.UserRequirement}}

## OUTPUT INSTRUCTIONS
1. Create a VBA script that generates a professional report based on the requirements
2. Include options for formatting, headers, and totals
3. Consider adding charts if appropriate for the data
4. Create the report in a new worksheet
5. Add proper error handling
6. Make the report visually appealing and easy to understand
7. Return only the VBA code, without additional explanations
`

// dataProcessingTemplate is specialized for data processing tasks
const dataProcessingTemplate = `# TASK: Generate Excel VBA data processing script

## SYSTEM INFORMATION
- Framework: exaMCP (Excel Automation with Model Context Protocol)
- Timestamp: {{.CurrentDateTime}}
- Target Excel Version: {{.Config.TargetExcelVersion}}
- Output Type: Data Processing

## SOURCE DATA
- Sheet: {{.Structure.SheetName}}
- Range: {{.Structure.RangeAddress}}
- Total Rows: {{.Structure.DataRows}}
- Has Headers: {{.Structure.HasHeaders}}
- Description: {{.Structure.Description}}

## DATA COLUMNS
{{.HeadersFormatted}}

## SAMPLE DATA
{{range $index, $row := .SampleDataLimited}}Row {{add $index 1}}: {{join $row ", "}}
{{end}}

{{if .ModuleDescriptions}}
## STANDARD MODULES AVAILABLE
{{.ModuleDescriptions}}
{{end}}

{{if .Config.IncludeExamples}}
## EXAMPLES
{{.Examples}}
{{end}}

## DATA PROCESSING REQUIREMENTS
{{.UserRequirement}}

## OUTPUT INSTRUCTIONS
1. Create a VBA script that processes the data according to the requirements
2. Focus on efficiency and accuracy in data transformation
3. Validate input data before processing
4. Place processed data in a new worksheet
5. Add comprehensive error handling
6. Include progress indicators for long-running operations
7. Return only the VBA code, without additional explanations
`

// userInterfaceTemplate is specialized for creating UIs in Excel
const userInterfaceTemplate = `# TASK: Generate Excel VBA user interface script

## SYSTEM INFORMATION
- Framework: exaMCP (Excel Automation with Model Context Protocol)
- Timestamp: {{.CurrentDateTime}}
- Target Excel Version: {{.Config.TargetExcelVersion}}
- Output Type: User Interface

## CONNECTED DATA
- Sheet: {{.Structure.SheetName}}
- Range: {{.Structure.RangeAddress}}
- Total Rows: {{.Structure.DataRows}}
- Has Headers: {{.Structure.HasHeaders}}
- Description: {{.Structure.Description}}

## DATA FIELDS
{{.HeadersFormatted}}

## SAMPLE DATA
{{range $index, $row := .SampleDataLimited}}Row {{add $index 1}}: {{join $row ", "}}
{{end}}

{{if .ModuleDescriptions}}
## STANDARD MODULES AVAILABLE
{{.ModuleDescriptions}}
{{end}}

{{if .Config.IncludeExamples}}
## EXAMPLES
{{.Examples}}
{{end}}

## UI REQUIREMENTS
{{.UserRequirement}}

## OUTPUT INSTRUCTIONS
1. Create a VBA script that builds a user-friendly interface
2. Include appropriate controls (forms, buttons, etc.) based on requirements
3. Connect the UI to the data source
4. Implement input validation and user feedback
5. Make the interface intuitive and professional
6. Include error handling for all user interactions
7. Return only the VBA code, without additional explanations
`

// Example code snippets
const basicExample = `## Basic Example
User Requirement: Filter records where [Sales] > 1000 and calculate the sum of [Quantity]

\`\`\`vba
Sub Main()
    On Error GoTo ErrorHandler
    
    ' Create result sheet
    Dim resultSheet As Worksheet
    Set resultSheet = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
    resultSheet.Name = "Result_" & Format(Now(), "yyyymmdd_hhnnss")
    
    ' Declare variables
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim cell As Range
    Dim salesCol As Integer, quantityCol As Integer
    Dim totalQuantity As Double
    
    ' Set worksheet and data range
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Set dataRange = ws.Range("A1:E100")
    
    ' Find column indices
    Dim headerRow As Range
    Set headerRow = dataRange.Rows(1)
    
    salesCol = 0
    quantityCol = 0
    
    For i = 1 To headerRow.Cells.Count
        If headerRow.Cells(i).Value = "Sales" Then
            salesCol = i
        ElseIf headerRow.Cells(i).Value = "Quantity" Then
            quantityCol = i
        End If
    Next i
    
    ' Validate columns were found
    If salesCol = 0 Or quantityCol = 0 Then
        MsgBox "Required columns not found", vbExclamation
        Exit Sub
    End If
    
    ' Create result header
    resultSheet.Range("A1").Value = "Filter Criteria: Sales > 1000"
    resultSheet.Range("A3").Value = "Total Quantity:"
    
    ' Process data and calculate
    totalQuantity = 0
    
    For i = 2 To dataRange.Rows.Count ' Skip header row
        If dataRange.Cells(i, salesCol).Value > 1000 Then
            totalQuantity = totalQuantity + dataRange.Cells(i, quantityCol).Value
        End If
    Next i
    
    ' Output result
    resultSheet.Range("B3").Value = totalQuantity
    
    ' Format result
    resultSheet.Range("A1").Font.Bold = True
    resultSheet.Range("A3:B3").Font.Bold = True
    
    MsgBox "Processing complete!", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
\`\`\`
`

const sqlExample = `## SQL Example
User Requirement: Use SQL to group by [Region] and calculate total sales and average order amount

\`\`\`vba
Sub Main()
    On Error GoTo ErrorHandler
    
    ' Create result sheet
    Dim resultSheet As Worksheet
    Set resultSheet = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
    resultSheet.Name = "RegionStats"
    
    ' SQL query
    Dim sqlQuery As String
    sqlQuery = "SELECT [Region], SUM([Sales]) AS TotalSales, AVG([OrderAmount]) AS AvgOrder " & _
              "FROM [Sheet1$] " & _
              "GROUP BY [Region] " & _
              "ORDER BY SUM([Sales]) DESC"
    
    ' Execute SQL query
    Call getSQL(sqlQuery, resultSheet.Range("A1"), True)
    
    ' Format result table
    With resultSheet.Range("A1").CurrentRegion
        ' Add title
        resultSheet.Range("A1:C1").Font.Bold = True
        
        ' Format numeric columns
        .Columns(2).NumberFormat = "#,##0.00"
        .Columns(3).NumberFormat = "#,##0.00"
        
        ' Auto-fit columns
        .EntireColumn.AutoFit
    End With
    
    MsgBox "Region statistics complete!", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
\`\`\`
`

const reportingExample = `## Reporting Example
User Requirement: Create a monthly sales report with a chart showing trends

\`\`\`vba
Sub CreateMonthlyReport()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' Create new report sheet
    Dim reportSheet As Worksheet
    Set reportSheet = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
    reportSheet.Name = "Monthly_Report_" & Format(Now(), "yyyymm")
    
    ' Variables
    Dim dataSheet As Worksheet
    Dim dataRange As Range
    Dim lastRow As Long, lastCol As Long
    Dim headerRow As Range
    Dim dateCol As Integer, salesCol As Integer, productCol As Integer
    Dim summaryTable As Range
    
    ' Set data source
    Set dataSheet = ThisWorkbook.Sheets("Sales")
    lastRow = dataSheet.Cells(dataSheet.Rows.Count, "A").End(xlUp).Row
    lastCol = dataSheet.Cells(1, dataSheet.Columns.Count).End(xlToLeft).Column
    Set dataRange = dataSheet.Range(dataSheet.Cells(1, 1), dataSheet.Cells(lastRow, lastCol))
    
    ' Find columns
    Set headerRow = dataRange.Rows(1)
    For i = 1 To headerRow.Columns.Count
        Select Case headerRow.Cells(1, i).Value
            Case "Date"
                dateCol = i
            Case "Sales"
                salesCol = i
            Case "Product"
                productCol = i
        End Select
    Next i
    
    ' Add report title
    With reportSheet
        .Range("A1").Value = "Monthly Sales Report"
        .Range("A2").Value = "Generated: " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
        .Range("A1").Font.Size = 16
        .Range("A1:A2").Font.Bold = True
        .Range("A4").Value = "Summary by Month"
    End With
    
    ' Create SQL query for monthly summary
    Dim sqlQuery As String
    sqlQuery = "SELECT Format([Date], 'yyyy-mm') AS Month, " & _
              "SUM([Sales]) AS TotalSales, " & _
              "COUNT([Sales]) AS OrderCount, " & _
              "AVG([Sales]) AS AvgOrderSize " & _
              "FROM [" & dataSheet.Name & "$] " & _
              "GROUP BY Format([Date], 'yyyy-mm') " & _
              "ORDER BY Format([Date], 'yyyy-mm')"
    
    ' Execute query
    Call getSQL(sqlQuery, reportSheet.Range("A5"), True)
    
    ' Format summary table
    Set summaryTable = reportSheet.Range("A5").CurrentRegion
    With summaryTable
        .Borders.LineStyle = xlContinuous
        .Font.Size = 11
        .Rows(1).Font.Bold = True
        .Columns(2).NumberFormat = "$#,##0.00"
        .Columns(4).NumberFormat = "$#,##0.00"
        .EntireColumn.AutoFit
    End With
    
    ' Create chart
    Dim chartObj As ChartObject
    Dim chartData As Range
    
    Set chartData = summaryTable
    Set chartObj = reportSheet.ChartObjects.Add(Left:=reportSheet.Range("F5").Left, _
                                              Top:=reportSheet.Range("F5").Top, _
                                              Width:=450, _
                                              Height:=250)
    
    With chartObj.Chart
        .SetSourceData Source:=chartData
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Monthly Sales Trend"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Sales ($)"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Month"
        .HasLegend = False
    End With
    
    ' Product breakdown
    reportSheet.Range("A" & summaryTable.Rows.Count + 7).Value = "Sales by Product"
    
    Dim productSQL As String
    productSQL = "SELECT [Product], " & _
                "SUM([Sales]) AS TotalSales, " & _
                "COUNT([Sales]) AS OrderCount " & _
                "FROM [" & dataSheet.Name & "$] " & _
                "GROUP BY [Product] " & _
                "ORDER BY SUM([Sales]) DESC"
    
    Call getSQL(productSQL, reportSheet.Range("A" & summaryTable.Rows.Count + 8), True)
    
    ' Format product table
    Dim productTable As Range
    Set productTable = reportSheet.Range("A" & summaryTable.Rows.Count + 8).CurrentRegion
    With productTable
        .Borders.LineStyle = xlContinuous
        .Font.Size = 11
        .Rows(1).Font.Bold = True
        .Columns(2).NumberFormat = "$#,##0.00"
        .EntireColumn.AutoFit
    End With
    
    ' Create pie chart for product breakdown
    Dim pieChart As ChartObject
    Set pieChart = reportSheet.ChartObjects.Add(Left:=reportSheet.Range("F" & summaryTable.Rows.Count + 8).Left, _
                                              Top:=reportSheet.Range("F" & summaryTable.Rows.Count + 8).Top, _
                                              Width:=450, _
                                              Height:=250)
    
    With pieChart.Chart
        .SetSourceData Source:=productTable
        .ChartType = xlPie
        .HasTitle = True
        .ChartTitle.Text = "Sales by Product"
        .HasLegend = True
        .Legend.Position = xlLegendPositionRight
    End With
    
    Application.ScreenUpdating = True
    reportSheet.Activate
    MsgBox "Monthly sales report generated successfully!", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error generating report: " & Err.Description, vbCritical
End Sub
\`\`\`
`

const dataProcessingExample = `## Data Processing Example
User Requirement: Clean data by removing duplicates, formatting dates, and standardizing product names

\`\`\`vba
Sub CleanData()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Create new sheet for cleaned data
    Dim sourceSheet As Worksheet
    Dim cleanSheet As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim sourceRange As Range
    Dim headerRow As Range
    Dim dateCol As Integer, productCol As Integer
    
    ' Display progress
    Application.StatusBar = "Initializing data cleaning process..."
    
    ' Set source sheet
    Set sourceSheet = ThisWorkbook.Sheets("RawData")
    
    ' Check if cleaned data sheet exists, if so delete it
    On Error Resume Next
    Set cleanSheet = ThisWorkbook.Sheets("CleanedData")
    If Not cleanSheet Is Nothing Then
        Application.DisplayAlerts = False
        cleanSheet.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo ErrorHandler
    
    ' Create new sheet
    Set cleanSheet = ThisWorkbook.Sheets.Add(After:=sourceSheet)
    cleanSheet.Name = "CleanedData"
    
    ' Get data range
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).Row
    lastCol = sourceSheet.Cells(1, sourceSheet.Columns.Count).End(xlToLeft).Column
    Set sourceRange = sourceSheet.Range(sourceSheet.Cells(1, 1), sourceSheet.Cells(lastRow, lastCol))
    
    ' Copy data to new sheet for processing
    sourceRange.Copy cleanSheet.Range("A1")
    
    ' Find important columns
    Set headerRow = cleanSheet.Range("1:1")
    For i = 1 To headerRow.Columns.Count
        Select Case headerRow.Cells(1, i).Value
            Case "Date", "OrderDate", "TransactionDate"
                dateCol = i
            Case "Product", "ProductName", "Item"
                productCol = i
        End Select
    Next i
    
    ' Update status
    Application.StatusBar = "Formatting dates..."
    
    ' Format dates
    If dateCol > 0 Then
        Dim dateRange As Range
        Set dateRange = cleanSheet.Range(cleanSheet.Cells(2, dateCol), cleanSheet.Cells(lastRow, dateCol))
        
        For Each cell In dateRange
            If Not IsEmpty(cell) Then
                If IsDate(cell.Value) Then
                    cell.NumberFormat = "yyyy-mm-dd"
                    cell.Value = DateValue(cell.Value)
                Else
                    ' Try to fix common date format issues
                    If Len(cell.Value) = 8 And IsNumeric(cell.Value) Then
                        ' YYYYMMDD format
                        cell.Value = DateSerial(Left(cell.Value, 4), Mid(cell.Value, 5, 2), Right(cell.Value, 2))
                        cell.NumberFormat = "yyyy-mm-dd"
                    Else
                        cell.Interior.Color = RGB(255, 255, 0) ' Highlight problematic cells
                    End If
                End If
            End If
        Next cell
    End If
    
    ' Update status
    Application.StatusBar = "Standardizing product names..."
    
    ' Standardize product names
    If productCol > 0 Then
        Dim productList As Object
        Set productList = CreateObject("Scripting.Dictionary")
        Dim standardizedNames As Object
        Set standardizedNames = CreateObject("Scripting.Dictionary")
        
        ' Define standard replacements
        standardizedNames.Add "LAPTOP", "Laptop"
        standardizedNames.Add "DESKTOP", "Desktop"
        standardizedNames.Add "TABLET", "Tablet"
        standardizedNames.Add "MONITOR", "Monitor"
        standardizedNames.Add "KEYBOARD", "Keyboard"
        standardizedNames.Add "MOUSE", "Mouse"
        
        ' Process product names
        Dim productRange As Range
        Set productRange = cleanSheet.Range(cleanSheet.Cells(2, productCol), cleanSheet.Cells(lastRow, productCol))
        
        For Each cell In productRange
            If Not IsEmpty(cell) Then
                ' Trim whitespace
                cell.Value = Trim(cell.Value)
                
                ' Convert standard names
                Dim productName As String
                productName = cell.Value
                
                ' Check for known replacements
                For Each key In standardizedNames.Keys
                    If InStr(1, UCase(productName), key, vbTextCompare) > 0 Then
                        productName = Replace(productName, key, standardizedNames(key), 1, -1, vbTextCompare)
                    End If
                Next
                
                cell.Value = productName
            End If
        Next cell
    End If
    
    ' Update status
    Application.StatusBar = "Removing duplicates..."
    
    ' Remove duplicates
    On Error Resume Next
    cleanSheet.Range("A1").CurrentRegion.RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5), Header:=xlYes
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "Could not automatically remove duplicates. They might need manual review.", vbInformation
    End If
    On Error GoTo ErrorHandler
    
    ' Format as table
    cleanSheet.Range("A1").CurrentRegion.Select
    cleanSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = "CleanData"
    
    ' Autofit columns
    cleanSheet.Cells.EntireColumn.AutoFit
    
    ' Reset application settings
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    cleanSheet.Activate
    MsgBox "Data cleaning complete!" & vbNewLine & _
           "Please review any yellow highlighted cells for potential date issues.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    ' Reset application settings
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error during data cleaning: " & Err.Description, vbCritical
End Sub
\`\`\`
`

const userInterfaceExample = `## User Interface Example
User Requirement: Create a data entry form for adding new records to the table

\`\`\`vba
' In a standard module:
Sub ShowDataEntryForm()
    DataEntryForm.Show
End Sub

' In a UserForm named "DataEntryForm":
Option Explicit

Private Sub UserForm_Initialize()
    ' Set form caption
    Me.Caption = "Data Entry Form"
    
    ' Initialize dropdown lists
    FillProductDropdown
    FillRegionDropdown
    
    ' Set default date to today
    txtDate.Value = Format(Date, "yyyy-mm-dd")
    
    ' Clear any previous values
    txtQuantity.Value = ""
    txtPrice.Value = ""
    
    ' Focus first field
    cboProduct.SetFocus
End Sub

Private Sub FillProductDropdown()
    ' Get unique product values from the data sheet
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim cell As Range
    Dim uniqueProducts As Object
    
    Set uniqueProducts = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Sheets("ProductData")
    
    ' Find product column
    Dim productCol As Integer
    For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If ws.Cells(1, i).Value = "Product" Then
            productCol = i
            Exit For
        End If
    Next i
    
    If productCol = 0 Then
        MsgBox "Product column not found!", vbExclamation
        Exit Sub
    End If
    
    ' Get last data row
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, productCol).End(xlUp).Row
    
    ' Build unique product list
    For Each cell In ws.Range(ws.Cells(2, productCol), ws.Cells(lastRow, productCol))
        If Not IsEmpty(cell.Value) And Not uniqueProducts.Exists(cell.Value) Then
            uniqueProducts.Add cell.Value, 1
            cboProduct.AddItem cell.Value
        End If
    Next cell
    
    ' If we have products, select the first one
    If cboProduct.ListCount > 0 Then
        cboProduct.ListIndex = 0
    End If
End Sub

Private Sub FillRegionDropdown()
    ' Add regions
    cboRegion.Clear
    cboRegion.AddItem "North"
    cboRegion.AddItem "South"
    cboRegion.AddItem "East"
    cboRegion.AddItem "West"
    cboRegion.AddItem "Central"
    
    ' Default to first region
    If cboRegion.ListCount > 0 Then
        cboRegion.ListIndex = 0
    End If
End Sub

Private Sub btnSave_Click()
    ' Validate form
    If Not ValidateForm Then
        Exit Sub
    End If
    
    ' Save data
    If SaveRecord Then
        MsgBox "Record saved successfully!", vbInformation
        
        ' Ask user if they want to enter another record
        If MsgBox("Do you want to enter another record?", vbQuestion + vbYesNo) = vbYes Then
            ' Clear form for new entry
            cboProduct.ListIndex = 0
            txtQuantity.Value = ""
            txtPrice.Value = ""
            txtDate.Value = Format(Date, "yyyy-mm-dd")
            cboProduct.SetFocus
        Else
            ' Close form
            Unload Me
        End If
    End If
End Sub

Private Function ValidateForm() As Boolean
    ' Check product
    If cboProduct.ListIndex = -1 Then
        MsgBox "Please select a product", vbExclamation
        cboProduct.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    ' Check region
    If cboRegion.ListIndex = -1 Then
        MsgBox "Please select a region", vbExclamation
        cboRegion.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    ' Check quantity
    If txtQuantity.Value = "" Then
        MsgBox "Please enter a quantity", vbExclamation
        txtQuantity.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    If Not IsNumeric(txtQuantity.Value) Then
        MsgBox "Quantity must be a number", vbExclamation
        txtQuantity.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    If Int(txtQuantity.Value) <> txtQuantity.Value Or txtQuantity.Value < 1 Then
        MsgBox "Quantity must be a positive whole number", vbExclamation
        txtQuantity.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    ' Check price
    If txtPrice.Value = "" Then
        MsgBox "Please enter a price", vbExclamation
        txtPrice.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    If Not IsNumeric(txtPrice.Value) Then
        MsgBox "Price must be a number", vbExclamation
        txtPrice.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    If CDbl(txtPrice.Value) <= 0 Then
        MsgBox "Price must be greater than zero", vbExclamation
        txtPrice.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    ' Check date
    If txtDate.Value = "" Then
        MsgBox "Please enter a date", vbExclamation
        txtDate.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    If Not IsDate(txtDate.Value) Then
        MsgBox "Please enter a valid date (yyyy-mm-dd)", vbExclamation
        txtDate.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    ' All validations passed
    ValidateForm = True
End Function

Private Function SaveRecord() As Boolean
    On Error GoTo ErrorHandler
    
    ' Get target worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("SalesData")
    
    ' Find last row
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Write data
    ws.Cells(lastRow, 1).Value = txtDate.Value
    ws.Cells(lastRow, 2).Value = cboProduct.Value
    ws.Cells(lastRow, 3).Value = cboRegion.Value
    ws.Cells(lastRow, 4).Value = txtQuantity.Value
    ws.Cells(lastRow, 5).Value = txtPrice.Value
    ws.Cells(lastRow, 6).Value = CDbl(txtQuantity.Value) * CDbl(txtPrice.Value)
    ws.Cells(lastRow, 7).Value = Now() ' Timestamp
    
    ' Format date cell
    ws.Cells(lastRow, 1).NumberFormat = "yyyy-mm-dd"
    
    ' Format numeric cells
    ws.Cells(lastRow, 4).NumberFormat = "0"
    ws.Cells(lastRow, 5).NumberFormat = "#,##0.00"
    ws.Cells(lastRow, 6).NumberFormat = "#,##0.00"
    
    SaveRecord = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error saving record: " & Err.Description, vbCritical
    SaveRecord = False
End Function

Private Sub btnCancel_Click()
    ' Close the form
    Unload Me
End Sub

Private Sub txtDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Validate and format date
    If txtDate.Value <> "" Then
        If IsDate(txtDate.Value) Then
            txtDate.Value = Format(CDate(txtDate.Value), "yyyy-mm-dd")
        End If
    End If
End Sub

Private Sub txtPrice_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Allow only numbers and decimal point
    Select Case KeyAscii
        Case 48 To 57 ' 0-9
            ' Allow
        Case 46 ' Decimal point
            ' Check if already contains a decimal point
            If InStr(1, txtPrice.Value, ".") > 0 Then
                KeyAscii = 0 ' Cancel the keypress
            End If
        Case 8 ' Backspace
            ' Allow
        Case Else
            KeyAscii = 0 ' Cancel the keypress
    End Select
End Sub

Private Sub txtQuantity_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Allow only numbers
    Select Case KeyAscii
        Case 48 To 57 ' 0-9
            ' Allow
        Case 8 ' Backspace
            ' Allow
        Case Else
            KeyAscii = 0 ' Cancel the keypress
    End Select
End Sub
\`\`\`
`
