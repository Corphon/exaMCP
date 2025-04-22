package mcp

import (
	"bytes"
	"fmt"
	"regexp"
	"strings"
	"text/template"
	"time"
)

// AdvancedPromptConfig contains enhanced configuration for advanced prompt generation
type AdvancedPromptConfig struct {
	// Core settings
	Language            string            // Prompt language (default: "en")
	DetailLevel         string            // "Basic", "Intermediate", "Advanced"
	FewShotExamples     int               // Number of examples to include (0-3)
	MaxSampleRows       int               // Maximum sample data rows to include
	
	// Task-specific settings
	TaskType            string            // Auto-detected or specified task type
	TargetExcelVersion  string            // Target Excel version
	
	// Contextual enhancements
	UserInfo            UserInfo          // Information about the current user
	HighlightColumns    []string          // Columns to emphasize in the prompt
	IncludeRelationships bool              // Include data relationships
	TemplateVariables   map[string]string // Custom template variables
	
	// Module options
	IncludeModules      []string          // Standard modules to include
	CustomModules       []CustomModule    // User-provided custom modules
	
	// Advanced features
	EnableChainOfThought bool              // Enable step-by-step reasoning
	IncludeErrorScenarios bool             // Include common error scenarios
	OptimizationLevel    string            // "None", "Basic", "Advanced"
}

// UserInfo contains information about the current user
type UserInfo struct {
	Username    string
	Timestamp   string
	SessionID   string
	Preferences map[string]string
}

// CustomModule represents a user-defined module to include in the prompt
type CustomModule struct {
	Name        string
	Description string
	CodeSample  string
}

// TaskClassification categorizes a user requirement
type TaskClassification struct {
	PrimaryType   string   // Main task type
	SecondaryType string   // Secondary task type
	Features      []string // Specific features required
	Complexity    string   // "Simple", "Moderate", "Complex"
	Keywords      []string // Key terms detected
}

// DefaultAdvancedConfig returns default configuration for advanced prompt generation
func DefaultAdvancedConfig() AdvancedPromptConfig {
	return AdvancedPromptConfig{
		Language:            "en",
		DetailLevel:         "Intermediate",
		FewShotExamples:     1,
		MaxSampleRows:       3,
		TaskType:            "Auto",
		TargetExcelVersion:  "Excel 2016+",
		IncludeRelationships: true,
		EnableChainOfThought: true,
		IncludeErrorScenarios: true,
		OptimizationLevel:    "Basic",
		TemplateVariables:    map[string]string{},
		HighlightColumns:     []string{},
		IncludeModules:       []string{},
		UserInfo: UserInfo{
			Username:    "User",
			Timestamp:   time.Now().Format("2006-01-02 15:04:05"),
			Preferences: map[string]string{},
		},
	}
}

// AdvancedMCPPrompt generates an enhanced MCP prompt with rich context and examples
func AdvancedMCPPrompt(structure DataRange, userRequirement string, includeStandardModules bool) string {
	// Get current time for timestamp
	currentTime := time.Now().UTC().Format("2006-01-02 15:04:05")
	
	// Initialize config with defaults and user data
	config := DefaultAdvancedConfig()
	config.UserInfo.Username = "Corphon"
	config.UserInfo.Timestamp = currentTime
	
	// Include standard modules if requested
	if includeStandardModules {
		config.IncludeModules = []string{"SQLUtils", "DataTools", "UIHelpers"}
	}
	
	// Auto-detect task type based on user requirement
	taskClassification := classifyUserRequirement(userRequirement, structure)
	config.TaskType = taskClassification.PrimaryType
	
	// Generate the enhanced prompt
	return GenerateAdvancedPrompt(structure, userRequirement, config)
}

// GenerateAdvancedPrompt creates a sophisticated prompt based on the provided configuration
func GenerateAdvancedPrompt(structure DataRange, userRequirement string, config AdvancedPromptConfig) string {
	var prompt strings.Builder

	// Select the appropriate template based on task type and detail level
	templateContent := selectPromptTemplate(config.TaskType, config.DetailLevel)

	// Prepare template data with rich context
	data := map[string]interface{}{
		"User":              config.UserInfo.Username,
		"Timestamp":         config.UserInfo.Timestamp,
		"Structure":         structure,
		"UserRequirement":   userRequirement,
		"Config":            config,
		"TaskClassification": classifyUserRequirement(userRequirement, structure),
		"HeadersFormatted":   formatHeadersAdvanced(structure.Headers, structure.DataTypes, config.HighlightColumns),
		"SampleData":         limitSampleData(structure.SampleData, config.MaxSampleRows),
		"RelationshipInfo":   getRelationshipDescription(structure.Relationships),
		"ModulesInfo":        getModulesDescription(config.IncludeModules),
		"Examples":           selectExamplesForTask(config.TaskType, config.FewShotExamples),
		"ChainOfThought":     getChainOfThoughtPrompt(config.TaskType),
		"ErrorScenarios":     getCommonErrorScenarios(config.TaskType),
		"OptimizationTips":   getOptimizationTips(config.OptimizationLevel),
	}

	// Add custom template variables
	for key, value := range config.TemplateVariables {
		data[key] = value
	}

	// Parse and execute template
	tmpl, err := template.New("advanced_mcp_prompt").Funcs(template.FuncMap{
		"add":  func(a, b int) int { return a + b },
		"join": strings.Join,
	}).Parse(templateContent)
	
	if err != nil {
		return fallbackAdvancedPrompt(structure, userRequirement, err)
	}

	var buf bytes.Buffer
	if err := tmpl.Execute(&buf, data); err != nil {
		return fallbackAdvancedPrompt(structure, userRequirement, err)
	}

	prompt.WriteString(buf.String())
	return prompt.String()
}

// classifyUserRequirement analyzes a user requirement to determine its type and complexity
func classifyUserRequirement(requirement string, structure DataRange) TaskClassification {
	classification := TaskClassification{
		PrimaryType:   "Generic",
		SecondaryType: "",
		Features:      []string{},
		Complexity:    "Moderate",
		Keywords:      []string{},
	}

	// Convert to lowercase for case-insensitive matching
	req := strings.ToLower(requirement)

	// Extract keywords
	keywords := extractKeywords(req)
	classification.Keywords = keywords

	// Task type detection patterns
	taskPatterns := map[string][]string{
		"Reporting": {
			"report", "chart", "graph", "visualize", "dashboard", "summary", "pivot",
			"statistics", "trend", "analysis", "kpi", "metric", "visualization",
		},
		"DataProcessing": {
			"filter", "sort", "clean", "calculate", "transform", "convert", "normalize",
			"aggregate", "group", "consolidate", "merge", "join", "data processing",
		},
		"UserInterface": {
			"form", "button", "userform", "interface", "ui", "input", "dialog",
			"dropdown", "checkbox", "interactive", "menu", "user experience",
		},
		"Automation": {
			"automate", "schedule", "batch", "periodic", "monitor", "workflow", 
			"trigger", "event", "automatic", "background", "routine",
		},
		"DataValidation": {
			"validate", "verify", "check", "ensure", "integrity", "constraint",
			"rule", "validation", "error checking", "valid", "format checking",
		},
	}

	// Detect primary task type
	primaryScore := make(map[string]int)
	for taskType, patterns := range taskPatterns {
		for _, pattern := range patterns {
			count := strings.Count(req, pattern)
			if count > 0 {
				primaryScore[taskType] += count
			}
		}
	}

	// Find the highest scoring task type
	highestScore := 0
	for taskType, score := range primaryScore {
		if score > highestScore {
			highestScore = score
			classification.PrimaryType = taskType
		}
	}

	// Detect secondary task type
	secondaryScore := make(map[string]int)
	for taskType, score := range primaryScore {
		if taskType != classification.PrimaryType && score > 0 {
			secondaryScore[taskType] = score
		}
	}

	// Find the highest scoring secondary task type
	highestSecondaryScore := 0
	for taskType, score := range secondaryScore {
		if score > highestSecondaryScore {
			highestSecondaryScore = score
			classification.SecondaryType = taskType
		}
	}

	// Detect features
	featurePatterns := map[string]string{
		"SQL":           "sql|query|select|from|where|group by",
		"Charts":        "chart|graph|plot|visualize|pie|bar|line",
		"Formatting":    "format|style|color|conditional|highlight",
		"ImportExport":  "import|export|csv|text file|external",
		"Calculations":  "calculate|sum|average|count|formula",
		"AdvancedUI":    "userform|complex form|multi-step|wizard",
		"ErrorHandling": "error handling|validation|try catch|on error",
	}

	for feature, pattern := range featurePatterns {
		if regexp.MustCompile(pattern).MatchString(req) {
			classification.Features = append(classification.Features, feature)
		}
	}

	// Determine complexity
	complexityScore := 0
	
	// Complexity factors
	if len(classification.Features) >= 3 {
		complexityScore += 2
	} else if len(classification.Features) >= 1 {
		complexityScore += 1
	}
	
	if classification.SecondaryType != "" {
		complexityScore += 1
	}
	
	if strings.Contains(req, "complex") || strings.Contains(req, "advanced") || 
	   strings.Contains(req, "sophisticated") {
		complexityScore += 1
	}
	
	if len(requirement) > 200 {
		complexityScore += 1
	}
	
	// Set complexity based on score
	if complexityScore >= 3 {
		classification.Complexity = "Complex"
	} else if complexityScore >= 1 {
		classification.Complexity = "Moderate"
	} else {
		classification.Complexity = "Simple"
	}

	return classification
}

// extractKeywords extracts important keywords from a requirement
func extractKeywords(requirement string) []string {
	// List of common technical terms in Excel automation
	technicalTerms := []string{
		"filter", "sort", "report", "chart", "data", "column", "row", "cell",
		"sheet", "workbook", "range", "formula", "function", "macro", "calculate",
		"format", "conditional", "pivot", "table", "validation", "userform",
		"button", "combobox", "textbox", "query", "sql", "export", "import",
		"sum", "average", "count", "unique", "duplicate", "error", "loop",
	}

	keywords := []string{}
	lowerReq := strings.ToLower(requirement)

	for _, term := range technicalTerms {
		if strings.Contains(lowerReq, term) {
			keywords = append(keywords, term)
		}
	}

	// Deduplicate
	uniqueKeywords := []string{}
	seen := make(map[string]bool)

	for _, kw := range keywords {
		if !seen[kw] {
			uniqueKeywords = append(uniqueKeywords, kw)
			seen[kw] = true
		}
	}

	// Limit to top 10 keywords
	if len(uniqueKeywords) > 10 {
		uniqueKeywords = uniqueKeywords[:10]
	}

	return uniqueKeywords
}

// selectPromptTemplate selects the most appropriate template based on task type and detail level
func selectPromptTemplate(taskType string, detailLevel string) string {
	// Select template based on task type
	var template string
	
	switch taskType {
	case "Reporting":
		template = advancedReportingTemplate
	case "DataProcessing":
		template = advancedDataProcessingTemplate
	case "UserInterface":
		template = advancedUserInterfaceTemplate
	case "Automation":
		template = advancedAutomationTemplate
	case "DataValidation":
		template = advancedDataValidationTemplate
	default:
		// Use generic template
		template = advancedGenericTemplate
	}
	
	// Further customize based on detail level
	if detailLevel == "Basic" {
		// Simplify template if needed
		template = strings.Replace(template, "## CHAIN OF THOUGHT\n{{.ChainOfThought}}\n\n", "", -1)
		template = strings.Replace(template, "## ERROR SCENARIOS\n{{.ErrorScenarios}}\n\n", "", -1)
		template = strings.Replace(template, "## OPTIMIZATION TIPS\n{{.OptimizationTips}}\n\n", "", -1)
	} else if detailLevel == "Advanced" {
		// Further enhance template if needed
		// (template already has advanced features)
	}
	
	return template
}

// formatHeadersAdvanced creates a detailed header section with enhanced formatting
func formatHeadersAdvanced(headers []string, dataTypes map[string]string, highlightColumns []string) string {
	var result strings.Builder
	
	for i, header := range headers {
		dataType := dataTypes[header]
		if dataType == "" {
			dataType = "Unknown"
		}
		
		// Check if this is a highlighted column
		isHighlighted := false
		for _, highlighted := range highlightColumns {
			if highlighted == header {
				isHighlighted = true
				break
			}
		}
		
		if isHighlighted {
			result.WriteString(fmt.Sprintf("[header:%s] (Column %s, Type: %s) *KEY COLUMN*\n", 
				header, columnLetterFromIndex(i), dataType))
		} else {
			result.WriteString(fmt.Sprintf("[header:%s] (Column %s, Type: %s)\n", 
				header, columnLetterFromIndex(i), dataType))
		}
	}
	
	return result.String()
}

// getRelationshipDescription generates a readable description of data relationships
func getRelationshipDescription(relationships []Relationship) string {
	if len(relationships) == 0 {
		return "No explicit relationships detected between data ranges."
	}
	
	var result strings.Builder
	
	result.WriteString("The following relationships exist between data elements:\n\n")
	
	for i, rel := range relationships {
		result.WriteString(fmt.Sprintf("%d. %s relationship: Field [%s] connects to [%s] in range %s\n", 
			i+1, rel.Type, rel.SourceField, rel.TargetField, rel.TargetRange))
	}
	
	return result.String()
}

// getModulesDescription creates descriptions of available modules
func getModulesDescription(modules []string) string {
	if len(modules) == 0 {
		return ""
	}
	
	var result strings.Builder
	
	for _, module := range modules {
		switch module {
		case "SQLUtils":
			result.WriteString("### SQLUtils Module\n")
			result.WriteString("A utility module for executing SQL queries against Excel data:\n\n")
			result.WriteString("```vba\n")
			result.WriteString("' Execute SQL query against Excel data and output results to a range\n")
			result.WriteString("Sub getSQL(StrSQL As String, rng As Range, Optional title As Boolean = True)\n")
			result.WriteString("    ' Uses ADO to query Excel data as a database\n")
			result.WriteString("    ' Parameters:\n")
			result.WriteString("    '   StrSQL - SQL query string (supports SELECT, GROUP BY, ORDER BY, etc.)\n")
			result.WriteString("    '   rng - Target range where results will be placed\n")
			result.WriteString("    '   title - Whether to include column headers (default: True)\n")
			result.WriteString("End Sub\n")
			result.WriteString("```\n\n")
			result.WriteString("Example usage:\n")
			result.WriteString("```vba\n")
			result.WriteString("Dim sql As String\n")
			result.WriteString("sql = \"SELECT [Region], SUM([Sales]) AS TotalSales FROM [Sheet1$] GROUP BY [Region] ORDER BY SUM([Sales]) DESC\"\n")
			result.WriteString("Call getSQL(sql, Worksheets(\"Report\").Range(\"A1\"), True)\n")
			result.WriteString("```\n\n")
			
		case "DataTools":
			result.WriteString("### DataTools Module\n")
			result.WriteString("A collection of functions for common data manipulation tasks:\n\n")
			result.WriteString("```vba\n")
			result.WriteString("' Find row number containing a value in a range\n")
			result.WriteString("Function FindRow(searchRange As Range, searchValue As Variant) As Long\n\n")
			result.WriteString("' Copy data between sheets with flexible options\n")
			result.WriteString("Sub CopyRangeToSheet(sourceRange As Range, targetSheet As Worksheet, targetCell As String)\n\n")
			result.WriteString("' Advanced sort for data ranges\n")
			result.WriteString("Sub SortRange(dataRange As Range, sortColumn As Integer, ascending As Boolean)\n\n")
			result.WriteString("' Remove duplicate values from a range\n")
			result.WriteString("Sub RemoveDuplicates(dataRange As Range, columnIndexes As Variant)\n")
			result.WriteString("```\n\n")
			
		case "UIHelpers":
			result.WriteString("### UIHelpers Module\n")
			result.WriteString("Utilities for creating user interfaces without building UserForms manually:\n\n")
			result.WriteString("```vba\n")
			result.WriteString("' Create a simple input form and return entered values\n")
			result.WriteString("Function CreateInputForm(title As String, fields As Variant) As Variant\n\n")
			result.WriteString("' Display a progress bar during long operations\n")
			result.WriteString("Sub ShowProgressBar(title As String, max As Long)\n")
			result.WriteString("Sub UpdateProgress(value As Long)\n")
			result.WriteString("Sub CloseProgressBar()\n\n")
			result.WriteString("' Create a message with timeout\n")
			result.WriteString("Sub TimedMessage(message As String, durationSeconds As Integer)\n")
			result.WriteString("```\n\n")
		}
	}
	
	return result.String()
}

// selectExamplesForTask returns appropriate examples for the given task type
func selectExamplesForTask(taskType string, numExamples int) string {
	if numExamples <= 0 {
		return ""
	}
	
	var examples []string
	
	// Add task-specific example first
	switch taskType {
	case "Reporting":
		examples = append(examples, reportingExample)
	case "DataProcessing":
		examples = append(examples, dataProcessingExample)
	case "UserInterface":
		examples = append(examples, userInterfaceExample)
	case "Automation":
		examples = append(examples, automationExample)
	case "DataValidation":
		examples = append(examples, dataValidationExample)
	default:
		examples = append(examples, basicExample)
	}
	
	// Add SQL example if we need more than one example
	if numExamples >= 2 {
		examples = append(examples, sqlExample)
	}
	
	// Add general purpose example if we need more
	if numExamples >= 3 {
		examples = append(examples, errorHandlingExample)
	}
	
	// Limit to requested number
	if len(examples) > numExamples {
		examples = examples[:numExamples]
	}
	
	// Format examples
	var result strings.Builder
	for i, example := range examples {
		result.WriteString(fmt.Sprintf("### Example %d\n%s\n\n", i+1, example))
	}
	
	return result.String()
}

// getChainOfThoughtPrompt generates step-by-step reasoning prompts for the specified task
func getChainOfThoughtPrompt(taskType string) string {
	switch taskType {
	case "Reporting":
		return `When creating a reporting script, think through these steps:
1. First, identify what key metrics need to be calculated from the data
2. Determine appropriate grouping and filtering criteria based on the requirements
3. Decide on the most effective data presentation format (tables, charts, or both)
4. Plan the layout and formatting of the report for readability
5. Consider adding summary statistics and headers/footers
6. Include error handling specific to data retrieval and calculation issues
7. Implement export/save/print functionality if needed`

	case "DataProcessing":
		return `When creating a data processing script, think through these steps:
1. First, validate the input data structure against expectations
2. Identify which transformations need to be applied to each column
3. Determine the logical order of operations for maximum efficiency
4. Plan for handling exceptions and edge cases in the data
5. Consider memory usage for large datasets
6. Include progress indicators for long-running operations
7. Validate the processed data before final output`

	case "UserInterface":
		return `When creating a user interface script, think through these steps:
1. First, identify all the input fields and controls needed
2. Design a logical flow and tab order for the form
3. Plan data validation for each input field
4. Determine appropriate default values
5. Create clear visual feedback for users
6. Handle both normal submission and cancellation scenarios
7. Ensure the UI remains responsive during processing
8. Add input validation and user feedback for errors`

	case "Automation":
		return `When creating an automation script, think through these steps:
1. First, identify all the steps that need to be automated
2. Determine dependencies between steps and optimal sequence
3. Plan for error recovery at each step to prevent partial completion
4. Add status updates or logging for monitoring
5. Consider performance optimizations for repetitive operations
6. Create safeguards against unintended consequences
7. Include cleanup operations to ensure the environment is left in a consistent state`

	case "DataValidation":
		return `When creating a data validation script, think through these steps:
1. First, identify all the validation rules that need to be applied
2. Determine the appropriate validation method for each rule
3. Plan how to collect and report validation errors
4. Consider the user experience when validation fails
5. Add suggestions for fixing validation errors where possible
6. Include summary statistics on validation results
7. Plan for partial acceptance of data with warnings vs. critical errors`

	default:
		return `When creating the VBA script, think through these steps:
1. First, understand exactly what data the script needs to work with
2. Break down the user requirement into discrete logical steps
3. Determine the most efficient approach for each step
4. Plan error handling for potential issues
5. Consider user experience and feedback
6. Ensure proper cleanup of resources
7. Add clear comments to explain the logic`
	}
}

// getCommonErrorScenarios provides examples of errors to handle for the specified task
func getCommonErrorScenarios(taskType string) string {
	// Basic error scenarios for all task types
	basic := `- Missing or invalid input data
- Required columns not found in the dataset
- Unexpected data types in cells
- Insufficient permissions to perform operations
- Out of memory for large datasets`
	
	// Task-specific error scenarios
	switch taskType {
	case "Reporting":
		return basic + `
- Division by zero in calculations
- Date range errors in time-based reports
- Chart creation fails due to invalid data
- Pivot table field references are invalid
- Report destination already exists or is locked`

	case "DataProcessing":
		return basic + `
- Text to number conversion errors
- Date parsing failures
- Duplicate key errors when consolidating data
- Formula calculation errors
- Target range not large enough for output data
- External data source connection failures`

	case "UserInterface":
		return basic + `
- Invalid user input formats
- Required fields left empty
- Inconsistent or conflicting selections
- Form canceled mid-operation
- Control array indexing errors
- Event handler errors during user interaction`

	case "Automation":
		return basic + `
- External application not available
- Operation timeout
- File access errors (locked, missing, corrupted)
- State inconsistency between operations
- Previously completed steps need to be undone after later failure
- Scheduled task conflicts`

	case "DataValidation":
		return basic + `
- Business rule violations in data
- Referential integrity issues
- Format validation failures
- Validation rules that conflict with each other
- Validation exceptions that need manual approval
- Missing validation reference data`

	default:
		return basic
	}
}

// getOptimizationTips provides performance optimization guidance for the specified level
func getOptimizationTips(level string) string {
	if level == "None" {
		return ""
	}
	
	// Basic optimization tips for all levels
	basic := `- Use Option Explicit to catch variable declaration errors
- Turn off screen updating, automatic calculation, and events during processing
- Use With blocks for repeated object references
- Minimize operations inside loops
- Declare appropriate variable types
- Read ranges into arrays for faster processing
- Write arrays back to ranges in one operation`
	
	if level == "Basic" {
		return basic
	}
	
	// Advanced optimization tips
	return basic + `
- Use early binding for external objects when possible
- Minimize Worksheet and Range object creation
- Release object references with Set obj = Nothing when finished
- Use For loops with explicit counters instead of For Each when possible
- Calculate ranges once and store in variables
- Use disconnected recordsets for complex data operations
- Add DoEvents in long-running processes to prevent Excel from appearing to hang
- Consider breaking very large operations into batched transactions
- Use error handling with resume capabilities for reliability
- Implement logging for diagnostics in complex scenarios`
}

// fallbackAdvancedPrompt provides a simpler prompt when the template system fails
func fallbackAdvancedPrompt(structure DataRange, userRequirement string, err error) string {
	var prompt strings.Builder
	
	prompt.WriteString("# TASK: Generate Excel VBA script based on user requirements\n\n")
	prompt.WriteString(fmt.Sprintf("ERROR: Advanced template processing failed: %v. Using fallback prompt.\n\n", err))
	
	// Include timestamp and user
	prompt.WriteString(fmt.Sprintf("Current Date and Time: %s\n", time.Now().UTC().Format("2006-01-02 15:04:05")))
	prompt.WriteString("User: Corphon\n\n")
	
	// Basic structure information
	prompt.WriteString("## EXCEL STRUCTURE\n")
	prompt.WriteString(fmt.Sprintf("- Sheet: %s\n", structure.SheetName))
	prompt.WriteString(fmt.Sprintf("- Range: %s\n", structure.RangeAddress))
	prompt.WriteString(fmt.Sprintf("- Headers: %s\n", strings.Join(structure.Headers, ", ")))
	prompt.WriteString(fmt.Sprintf("- Data Rows: %d\n\n", structure.DataRows))
	
	// Headers detail
	prompt.WriteString("## HEADERS\n")
	for i, header := range structure.Headers {
		dataType := structure.DataTypes[header]
		if dataType == "" {
			dataType = "Unknown"
		}
		prompt.WriteString(fmt.Sprintf("[header:%s] (Column %s, Type: %s)\n", 
			header, columnLetterFromIndex(i), dataType))
	}
	prompt.WriteString("\n")
	
	// Sample data
	prompt.WriteString("## SAMPLE DATA\n")
	for i, row := range structure.SampleData {
		if i >= 3 { // Limit to 3 rows
			break
		}
		prompt.WriteString(fmt.Sprintf("Row %d: %s\n", i+1, strings.Join(row, ", ")))
	}
	prompt.WriteString("\n")
	
	// User requirement
	prompt.WriteString("## USER REQUIREMENT\n")
	prompt.WriteString(userRequirement)
	prompt.WriteString("\n\n")
	
	// Output instructions
	prompt.WriteString("## OUTPUT INSTRUCTIONS\n")
	prompt.WriteString("Generate comprehensive VBA code that fulfills the user requirement.\n")
	prompt.WriteString("Include error handling and ensure the code is optimized for performance.\n")
	
	return prompt.String()
}

// Advanced template definitions

// advancedGenericTemplate is the base template for general tasks
const advancedGenericTemplate = `# TASK: Generate Excel VBA script based on user requirements

## SYSTEM INFORMATION
- Framework: exaMCP (Excel Automation with Model Context Protocol)
- Current Date and Time (UTC): {{.Timestamp}}
- User: {{.User}}
- Target Excel Version: {{.Config.TargetExcelVersion}}

## TASK ANALYSIS
- Primary Task Type: {{.TaskClassification.PrimaryType}}
{{if .TaskClassification.SecondaryType}}
- Secondary Task Type: {{.TaskClassification.SecondaryType}}
{{end}}
- Complexity Level: {{.TaskClassification.Complexity}}
- Key Features Required: {{range .TaskClassification.Features}}{{.}}, {{end}}

## EXCEL STRUCTURE
- Sheet: {{.Structure.SheetName}}
- Range: {{.Structure.RangeAddress}}
- Total Rows: {{.Structure.DataRows}}
- Has Headers: {{.Structure.HasHeaders}}
- Description: {{.Structure.Description}}

## HEADERS
{{.HeadersFormatted}}

## SAMPLE DATA
{{range $index, $row := .SampleData}}Row {{add $index 1}}: {{join $row ", "}}
{{end}}

{{if .RelationshipInfo}}
## DATA RELATIONSHIPS
{{.RelationshipInfo}}
{{end}}

{{if .ModulesInfo}}
## STANDARD MODULES AVAILABLE
{{.ModulesInfo}}
{{end}}

{{if .Examples}}
## EXAMPLES
{{.Examples}}
{{end}}

## CHAIN OF THOUGHT
{{.ChainOfThought}}

## ERROR SCENARIOS
Consider handling these common error cases:
{{.ErrorScenarios}}

## OPTIMIZATION TIPS
{{.OptimizationTips}}

## USER REQUIREMENT
{{.UserRequirement}}

## OUTPUT INSTRUCTIONS
1. Analyze the Excel structure and user requirement carefully
2. Generate complete, working VBA code that fulfills all requirements
3. Include proper error handling for robustness
4. Use descriptive variable names and add comments to explain complex logic
5. Utilize standard modules when appropriate for the task
6. Apply the optimization techniques mentioned above where relevant
7. Return only the VBA code, without additional explanations
`

// advancedReportingTemplate focuses on reporting and visualization
const advancedReportingTemplate = `# TASK: Generate Excel VBA reporting script

## SYSTEM INFORMATION
- Framework: exaMCP (Excel Automation with Model Context Protocol)
- Current Date and Time (UTC): {{.Timestamp}}
- User: {{.User}}
- Target Excel Version: {{.Config.TargetExcelVersion}}

## REPORTING TASK DETAILS
- Complexity Level: {{.TaskClassification.Complexity}}
- Secondary Aspects: {{if .TaskClassification.SecondaryType}}{{.TaskClassification.SecondaryType}}{{else}}None{{end}}
- Key Features Required: {{range .TaskClassification.Features}}{{.}}, {{end}}

## SOURCE DATA
- Sheet: {{.Structure.SheetName}}
- Range: {{.Structure.RangeAddress}}
- Total Rows: {{.Structure.DataRows}}
- Has Headers: {{.Structure.HasHeaders}}
- Description: {{.Structure.Description}}

## COLUMNS FOR REPORTING
{{.HeadersFormatted}}

## SAMPLE DATA
{{range $index, $row := .SampleData}}Row {{add $index 1}}: {{join $row ", "}}
{{end}}

{{if .RelationshipInfo}}
## DATA RELATIONSHIPS
{{.RelationshipInfo}}
{{end}}

{{if .ModulesInfo}}
## STANDARD MODULES AVAILABLE
{{.ModulesInfo}}
{{end}}

{{if .Examples}}
## EXAMPLES
{{.Examples}}
{{end}}

## REPORTING DESIGN CONSIDERATIONS
{{.ChainOfThought}}

## HANDLING REPORTING ERRORS
Consider handling these common reporting error scenarios:
{{.ErrorScenarios}}

## REPORT OPTIMIZATION TIPS
{{.OptimizationTips}}

## REPORTING REQUIREMENTS
{{.UserRequirement}}

## OUTPUT INSTRUCTIONS
1. Create a VBA script that generates a professional report based on the requirements
2. Include formatted headers, totals, and proper organization of information
3. Create appropriate visualizations (charts, conditional formatting) if relevant
4. Generate the report in a new worksheet with a descriptive name
5. Add export/print options if mentioned in requirements
6. Include robust error handling for all data operations
7. Format the output for professional presentation
8. Return only the VBA code, without additional explanations
`

// advancedDataProcessingTemplate focuses on data manipulation
const advancedDataProcessingTemplate = `# TASK: Generate Excel VBA data processing script

## SYSTEM INFORMATION
- Framework: exaMCP (Excel Automation with Model Context Protocol)
- Current Date and Time (UTC): {{.Timestamp}}
- User: {{.User}}
- Target Excel Version: {{.Config.TargetExcelVersion}}

## DATA PROCESSING TASK DETAILS
- Complexity Level: {{.TaskClassification.Complexity}}
- Secondary Aspects: {{if .TaskClassification.SecondaryType}}{{.TaskClassification.SecondaryType}}{{else}}None{{end}}
- Key Features Required: {{range .TaskClassification.Features}}{{.}}, {{end}}

## SOURCE DATA
- Sheet: {{.Structure.SheetName}}
- Range: {{.Structure.RangeAddress}}
- Total Rows: {{.Structure.DataRows}}
- Has Headers: {{.Structure.HasHeaders}}
- Description: {{.Structure.Description}}

## DATA COLUMNS
{{.HeadersFormatted}}

## SAMPLE DATA
{{range $index, $row := .SampleData}}Row {{add $index 1}}: {{join $row ", "}}
{{end}}

{{if .RelationshipInfo}}
## DATA RELATIONSHIPS
{{.RelationshipInfo}}
{{end}}

{{if .ModulesInfo}}
## STANDARD MODULES AVAILABLE
{{.ModulesInfo}}
{{end}}

{{if .Examples}}
## EXAMPLES
{{.Examples}}
{{end}}

## DATA PROCESSING APPROACH
{{.ChainOfThought}}

## DATA PROCESSING ERROR SCENARIOS
Consider handling these common data processing error cases:
{{.ErrorScenarios}}

## DATA PROCESSING OPTIMIZATION TIPS
{{.OptimizationTips}}

## DATA PROCESSING REQUIREMENTS
{{.UserRequirement}}

## OUTPUT INSTRUCTIONS
1. Create a VBA script that processes the data according to the requirements
2. Focus on data integrity, validation, and transformation accuracy
3. Implement efficient algorithms appropriate for the data volume
4. Include progress indicators for long-running operations
5. Place processed data in a well-structured output format
6. Add comprehensive error handling for all data operations
7. Validate results to ensure accuracy
8. Return only the VBA code, without additional explanations
`

// advancedUserInterfaceTemplate focuses on form creation and user interaction
const advancedUserInterfaceTemplate = `# TASK: Generate Excel VBA user interface script

## SYSTEM INFORMATION
- Framework: exaMCP (Excel Automation with Model Context Protocol)
- Current Date and Time (UTC): {{.Timestamp}}
- User: {{.User}}
- Target Excel Version: {{.Config.TargetExcelVersion}}

## UI TASK DETAILS
- Complexity Level: {{.TaskClassification.Complexity}}
- Secondary Aspects: {{if .TaskClassification.SecondaryType}}{{.TaskClassification.SecondaryType}}{{else}}None{{end}}
- Key Features Required: {{range .TaskClassification.Features}}{{.}}, {{end}}

## CONNECTED DATA
- Sheet: {{.Structure.SheetName}}
- Range: {{.Structure.RangeAddress}}
- Total Rows: {{.Structure.DataRows}}
- Has Headers: {{.Structure.HasHeaders}}
- Description: {{.Structure.Description}}

## DATA FIELDS
{{.HeadersFormatted}}

## SAMPLE DATA
{{range $index, $row := .SampleData}}Row {{add $index 1}}: {{join $row ", "}}
{{end}}

{{if .RelationshipInfo}}
## DATA RELATIONSHIPS
{{.RelationshipInfo}}
{{end}}

{{if .ModulesInfo}}
## STANDARD MODULES AVAILABLE
{{.ModulesInfo}}
{{end}}

{{if .Examples}}
## EXAMPLES
{{.Examples}}
{{end}}

## UI DESIGN APPROACH
{{.ChainOfThought}}

## UI ERROR SCENARIOS
Consider handling these common UI error cases:
{{.ErrorScenarios}}

## UI OPTIMIZATION TIPS
{{.OptimizationTips}}

## UI REQUIREMENTS
{{.UserRequirement}}

## OUTPUT INSTRUCTIONS
1. Create a VBA script that builds a user-friendly interface based on the requirements
2. Design appropriate forms with logical layout and professional appearance
3. Include all necessary controls with proper validation
4. Ensure the UI is intuitive and provides clear feedback to users
5. Add data binding between UI controls and Excel data
6. Implement proper event handling and form lifecycle management
7. Include error handling for all user interactions
8. Return only the VBA code, without additional explanations
`

// advancedAutomationTemplate focuses on process automation
const advancedAutomationTemplate = `# TASK: Generate Excel VBA automation script

## SYSTEM INFORMATION
- Framework: exaMCP (Excel Automation with Model Context Protocol)
- Current Date and Time (UTC): {{.Timestamp}}
- User: {{.User}}
- Target Excel Version: {{.Config.TargetExcelVersion}}

## AUTOMATION TASK DETAILS
- Complexity Level: {{.TaskClassification.Complexity}}
- Secondary Aspects: {{if .TaskClassification.SecondaryType}}{{.TaskClassification.SecondaryType}}{{else}}None{{end}}
- Key Features Required: {{range .TaskClassification.Features}}{{.}}, {{end}}

## CONNECTED DATA
- Sheet: {{.Structure.SheetName}}
- Range: {{.Structure.RangeAddress}}
- Total Rows: {{.Structure.DataRows}}
- Has Headers: {{.Structure.HasHeaders}}
- Description: {{.Structure.Description}}

## DATA FIELDS
{{.HeadersFormatted}}

## SAMPLE DATA
{{range $index, $row := .SampleData}}Row {{add $index 1}}: {{join $row ", "}}
{{end}}

{{if .RelationshipInfo}}
## DATA RELATIONSHIPS
{{.RelationshipInfo}}
{{end}}

{{if .ModulesInfo}}
## STANDARD MODULES AVAILABLE
{{.ModulesInfo}}
{{end}}

{{if .Examples}}
## EXAMPLES
{{.Examples}}
{{end}}

## AUTOMATION APPROACH
{{.ChainOfThought}}

## AUTOMATION ERROR SCENARIOS
Consider handling these common automation error cases:
{{.ErrorScenarios}}

## AUTOMATION OPTIMIZATION TIPS
{{.OptimizationTips}}

## AUTOMATION REQUIREMENTS
{{.UserRequirement}}

## OUTPUT INSTRUCTIONS
1. Create a VBA script that automates the required process
2. Design a reliable workflow with proper sequencing of operations
3. Include logging or status reporting for monitoring
4. Implement robust error handling with recovery mechanisms
5. Add safeguards against unintended data modification
6. Consider adding a user confirmation step before critical operations
7. Ensure the automation is efficient and reliable
8. Return only the VBA code, without additional explanations
`

// advancedDataValidationTemplate focuses on validating data
const advancedDataValidationTemplate = `# TASK: Generate Excel VBA data validation script

## SYSTEM INFORMATION
- Framework: exaMCP (Excel Automation with Model Context Protocol)
- Current Date and Time (UTC): {{.Timestamp}}
- User: {{.User}}
- Target Excel Version: {{.Config.TargetExcelVersion}}

## VALIDATION TASK DETAILS
- Complexity Level: {{.TaskClassification.Complexity}}
- Secondary Aspects: {{if .TaskClassification.SecondaryType}}{{.TaskClassification.SecondaryType}}{{else}}None{{end}}
- Key Features Required: {{range .TaskClassification.Features}}{{.}}, {{end}}

## DATA TO VALIDATE
- Sheet: {{.Structure.SheetName}}
- Range: {{.Structure.RangeAddress}}
- Total Rows: {{.Structure.DataRows}}
- Has Headers: {{.Structure.HasHeaders}}
- Description: {{.Structure.Description}}

## DATA FIELDS
{{.HeadersFormatted}}

## SAMPLE DATA
{{range $index, $row := .SampleData}}Row {{add $index 1}}: {{join $row ", "}}
{{end}}

{{if .RelationshipInfo}}
## DATA RELATIONSHIPS
{{.RelationshipInfo}}
{{end}}

{{if .ModulesInfo}}
## STANDARD MODULES AVAILABLE
{{.ModulesInfo}}
{{end}}

{{if .Examples}}
## EXAMPLES
{{.Examples}}
{{end}}

## VALIDATION APPROACH
{{.ChainOfThought}}

## VALIDATION ERROR SCENARIOS
Consider handling these common validation error cases:
{{.ErrorScenarios}}

## VALIDATION OPTIMIZATION TIPS
{{.OptimizationTips}}

## VALIDATION REQUIREMENTS
{{.UserRequirement}}

## OUTPUT INSTRUCTIONS
1. Create a VBA script that validates data according to the requirements
2. Implement all required validation rules with clear error reporting
3. Use appropriate validation methods (built-in Excel validation, custom logic)
4. Provide clear feedback about validation failures
5. Generate a validation summary report
6. Add options to highlight/mark invalid data
7. Include suggestions for fixing common validation issues
8. Return only the VBA code, without additional explanations
`

// Additional example: automation example
const automationExample = `## Automation Example
User Requirement: Automate the process of importing multiple CSV files, combining them into a single dataset, and creating a summary report

\`\`\`vba
Sub AutomateDataImport()
    On Error GoTo ErrorHandler
    
    ' Turn off screen updating for better performance
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Create a log sheet for tracking the process
    Dim logSheet As Worksheet
    On Error Resume Next
    Set logSheet = ThisWorkbook.Sheets("ImportLog")
    If logSheet Is Nothing Then
        Set logSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        logSheet.Name = "ImportLog"
    End If
    On Error GoTo ErrorHandler
    
    ' Initialize log
    logSheet.Cells.Clear
    logSheet.Range("A1").Value = "Import Process Log"
    logSheet.Range("A2").Value = "Started: " & Now()
    logSheet.Range("A4").Value = "File"
    logSheet.Range("B4").Value = "Status"
    logSheet.Range("C4").Value = "Records"
    logSheet.Range("D4").Value = "Timestamp"
    logSheet.Range("A1:D4").Font.Bold = True
    
    ' Create or clear the consolidated data sheet
    Dim dataSheet As Worksheet
    On Error Resume Next
    Set dataSheet = ThisWorkbook.Sheets("ConsolidatedData")
    If dataSheet Is Nothing Then
        Set dataSheet = ThisWorkbook.Sheets.Add(After:=logSheet)
        dataSheet.Name = "ConsolidatedData"
    Else
        dataSheet.Cells.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' Get the folder containing CSV files
    Dim folderPath As String
    folderPath = GetFolderPath()
    If folderPath = "" Then
        Application.StatusBar = False
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Exit Sub
    End If
    
    ' Log the selected folder
    logSheet.Range("A3").Value = "Folder: " & folderPath
    
    ' Initialize variables for tracking
    Dim totalFiles As Long, processedFiles As Long, totalRecords As Long
    Dim logRow As Long, dataRow As Long
    Dim hasHeaders As Boolean, firstFile As Boolean
    
    logRow = 5 ' Start logging from row 5
    dataRow = 1 ' Start data at row 1
    firstFile = True ' First file flag for headers
    hasHeaders = True ' Assume CSV files have headers
    
    ' Get list of CSV files
    Dim fileSystem As Object, folder As Object, file As Object, files As Object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set folder = fileSystem.GetFolder(folderPath)
    Set files = folder.Files
    
    ' Count CSV files
    totalFiles = 0
    For Each file In files
        If Right(LCase(file.Name), 4) = ".csv" Then
            totalFiles = totalFiles + 1
        End If
    Next file
    
    ' Process each CSV file
    processedFiles = 0
    For Each file In files
        ' Only process CSV files
        If Right(LCase(file.Name), 4) = ".csv" Then
            ' Update status
            processedFiles = processedFiles + 1
            Application.StatusBar = "Processing file " & processedFiles & " of " & totalFiles & ": " & file.Name
            
            ' Log the file
            logSheet.Range("A" & logRow).Value = file.Name
            logSheet.Range("D" & logRow).Value = Now()
            
            ' Import the CSV file
            Dim importSuccess As Boolean
            Dim recordCount As Long
            
            importSuccess = ImportCSVFile(file.Path, dataSheet, dataRow, firstFile, hasHeaders, recordCount)
            
            ' Update log
            If importSuccess Then
                logSheet.Range("B" & logRow).Value = "Success"
                logSheet.Range("C" & logRow).Value = recordCount
                totalRecords = totalRecords + recordCount
                
                ' Update data row counter for next file
                If firstFile Then
                    ' First file includes headers (if hasHeaders is True)
                    If hasHeaders Then
                        dataRow = dataRow + recordCount + 1
                    Else
                        dataRow = dataRow + recordCount
                    End If
                    firstFile = False
                Else
                    ' Subsequent files (skip headers if they have them)
                    dataRow = dataRow + recordCount
                End If
            Else
                logSheet.Range("B" & logRow).Value = "Failed"
                logSheet.Range("B" & logRow).Interior.Color = RGB(255, 200, 200)
            End If
            
            logRow = logRow + 1
        End If
    Next file
    
    ' Format the consolidated data as a table
    If dataRow > 1 Then
        Dim headerRow As Long
        If hasHeaders Then
            headerRow = 1
        Else
            headerRow = 0
        End If
        
        If headerRow > 0 Then
            Dim dataRange As Range
            Set dataRange = dataSheet.Range("A1").CurrentRegion
            
            ' Create a table
            Dim dataTable As ListObject
            On Error Resume Next
            Set dataTable = dataSheet.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
            If Not dataTable Is Nothing Then
                dataTable.Name = "ConsolidatedDataTable"
                dataTable.TableStyle = "TableStyleMedium2"
            End If
            On Error GoTo ErrorHandler
        End If
    End If
    
    ' Create summary report
    CreateSummaryReport totalFiles, processedFiles, totalRecords
    
    ' Clean up
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Final log entry
    logSheet.Range("A" & logRow).Value = "Import Completed"
    logSheet.Range("B" & logRow).Value = "Total Files: " & processedFiles
    logSheet.Range("C" & logRow).Value = "Total Records: " & totalRecords
    logSheet.Range("D" & logRow).Value = Now()
    logSheet.Range("A" & logRow & ":D" & logRow).Font.Bold = True
    
    ' Format log sheet
    logSheet.Columns("A:D").AutoFit
    logSheet.Activate
    
    MsgBox "Import process completed." & vbNewLine & _
           "Files processed: " & processedFiles & " of " & totalFiles & vbNewLine & _
           "Total records imported: " & totalRecords, vbInformation
    
    Exit Sub
    
ErrorHandler:
    ' Clean up in case of error
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Log the error
    If logSheet Is Nothing Then
        On Error Resume Next
        Set logSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        logSheet.Name = "ImportLog"
        logSheet.Range("A1").Value = "Import Process Log"
    End If
    
    On Error Resume Next
    logSheet.Range("A" & logRow).Value = "ERROR"
    logSheet.Range("B" & logRow).Value = Err.Description
    logSheet.Range("D" & logRow).Value = Now()
    logSheet.Range("A" & logRow & ":D" & logRow).Interior.Color = RGB(255, 150, 150)
    
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

' Function to get folder path from user
Function GetFolderPath() As String
    Dim folderDialog As Object
    Set folderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    With folderDialog
        .Title = "Select Folder Containing CSV Files"
        .AllowMultiSelect = False
        If .Show = -1 Then
            GetFolderPath = .SelectedItems(1)
        Else
            GetFolderPath = ""
        End If
    End With
End Function

' Function to import a CSV file
Function ImportCSVFile(filePath As String, targetSheet As Worksheet, startRow As Long, _
                       isFirstFile As Boolean, hasHeaders As Boolean, ByRef recordCount As Long) As Boolean
    On Error GoTo ImportError
    
    ' Set up QueryTable to import the CSV
    Dim qt As QueryTable
    Dim targetRange As Range
    Dim tempSheet As Worksheet
    
    ' Create a temporary sheet for import
    Set tempSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    tempSheet.Name = "TempImport_" & Format(Now(), "hhmmss")
    
    ' Set up QueryTable for CSV import
    Set targetRange = tempSheet.Range("A1")
    Set qt = tempSheet.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=targetRange)
    
    With qt
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(xlGeneralFormat, xlGeneralFormat, xlGeneralFormat, xlGeneralFormat, xlGeneralFormat, xlGeneralFormat, xlGeneralFormat, xlGeneralFormat, xlGeneralFormat, xlGeneralFormat)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
        .Delete
    End With
    
    ' Count imported records
    Dim usedRange As Range
    Set usedRange = tempSheet.UsedRange
    
    If usedRange.Rows.Count = 1 And Application.CountA(usedRange) = 0 Then
        ' Empty file
        recordCount = 0
        tempSheet.Delete
        ImportCSVFile = True
        Exit Function
    End If
    
    recordCount = usedRange.Rows.Count
    If hasHeaders Then
        recordCount = recordCount - 1
    End If
    
    ' Copy data to the consolidated sheet
    If isFirstFile Then
        ' First file - include everything
        usedRange.Copy targetSheet.Range("A" & startRow)
    Else
        ' Subsequent files - skip header row if exists
        If hasHeaders Then
            tempSheet.Range("A2:" & RangeColumn(usedRange.Columns.Count) & usedRange.Rows.Count).Copy _
                targetSheet.Range("A" & startRow)
        Else
            usedRange.Copy targetSheet.Range("A" & startRow)
        End If
    End If
    
    ' Delete temporary sheet
    Application.DisplayAlerts = False
    tempSheet.Delete
    Application.DisplayAlerts = True
    
    ImportCSVFile = True
    Exit Function
    
ImportError:
    ' Clean up on error
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not tempSheet Is Nothing Then tempSheet.Delete
    Application.DisplayAlerts = True
    
    recordCount = 0
    ImportCSVFile = False
End Function

' Function to get column letter from number
Function RangeColumn(colNum As Integer) As String
    If colNum <= 26 Then
        RangeColumn = Chr(64 + colNum)
    Else
        RangeColumn = Chr(Int((colNum - 1) / 26) + 64) & Chr(((colNum - 1) Mod 26) + 65)
    End If
End Function

' Procedure to create a summary report
Sub CreateSummaryReport(totalFiles As Long, processedFiles As Long, totalRecords As Long)
    On Error Resume Next
    
    ' Create or get summary sheet
    Dim summarySheet As Worksheet
    Set summarySheet = ThisWorkbook.Sheets("ImportSummary")
    If summarySheet Is Nothing Then
        Set summarySheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(1))
        summarySheet.Name = "ImportSummary"
    End If
    summarySheet.Cells.Clear
    
    ' Add summary information
    With summarySheet
        .Range("A1").Value = "Import Summary Report"
        .Range("A1").Font.Size = 14
        .Range("A1").Font.Bold = True
        
        .Range("A3").Value = "Date:"
        .Range("B3").Value = Date
        .Range("A4").Value = "Time:"
        .Range("B4").Value = Time
        
        .Range("A6").Value = "Total CSV Files:"
        .Range("B6").Value = totalFiles
        .Range("A7").Value = "Files Processed:"
        .Range("B7").Value = processedFiles
        .Range("A8").Value = "Success Rate:"
        If totalFiles > 0 Then
            .Range("B8").Value = Format(processedFiles / totalFiles, "0.0%")
        Else
            .Range("B8").Value = "N/A"
        End If
        
        .Range("A10").Value = "Total Records Imported:"
        .Range("B10").Value = totalRecords
        
        ' Add consolidated data statistics if available
        If ThisWorkbook.Sheets("ConsolidatedData").UsedRange.Rows.Count > 1 Then
            Dim dataSheet As Worksheet
            Set dataSheet = ThisWorkbook.Sheets("ConsolidatedData")
            
            ' Get column count for headers
            Dim headerCount As Integer
            headerCount = dataSheet.UsedRange.Columns.Count
            
            .Range("A12").Value = "Data Statistics:"
            .Range("A13").Value = "Columns:"
            .Range("B13").Value = headerCount
            
            ' List headers
            .Range("A15").Value = "Column Headers:"
            For i = 1 To headerCount
                .Cells(16, i).Value = dataSheet.Cells(1, i).Value
            Next i
            
            ' Format header list
            .Range(.Cells(16, 1), .Cells(16, headerCount)).Font.Bold = True
            .Range(.Cells(16, 1), .Cells(16, headerCount)).Borders.Weight = xlThin
        End If
        
        ' Format report
        .Columns("A:B").AutoFit
    End With
End Sub
\`\`\`
`

// Data validation example
const dataValidationExample = `## Data Validation Example
User Requirement: Create a data validation script that checks for duplicate order IDs, ensures dates are within the current quarter, and validates that all required fields are filled out

\`\`\`vba
Sub ValidateOrderData()
    On Error GoTo ErrorHandler
    
    ' Turn off screen updating for better performance
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Initialize variables
    Dim ws As Worksheet
    Dim resultSheet As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim headerRow As Range
    Dim validationResults As Object
    Set validationResults = CreateObject("Scripting.Dictionary")
    
    ' Get the data sheet
    On Error Resume Next
    Set ws = ActiveSheet
    If ws Is Nothing Then
        MsgBox "Please select a worksheet with order data.", vbExclamation
        GoTo CleanExit
    End If
    On Error GoTo ErrorHandler
    
    ' Create or clear validation results sheet
    On Error Resume Next
    Set resultSheet = ThisWorkbook.Sheets("ValidationResults")
    If resultSheet Is Nothing Then
        Set resultSheet = ThisWorkbook.Sheets.Add(After:=ws)
        resultSheet.Name = "ValidationResults"
    Else
        resultSheet.Cells.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' Find the data range
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow <= 1 Then
        MsgBox "No data found in the selected worksheet.", vbExclamation
        GoTo CleanExit
    End If
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Identify required columns
    Dim orderIdCol As Integer, orderDateCol As Integer, customerCol As Integer
    Dim productCol As Integer, quantityCol As Integer, priceCol As Integer
    Dim requiredCols As Object
    Set requiredCols = CreateObject("Scripting.Dictionary")
    
    ' Map column names to column indices
    Set headerRow = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
    
    For i = 1 To lastCol
        Select Case Trim(LCase(headerRow.Cells(1, i).Value))
            Case "order id", "orderid", "order_id"
                orderIdCol = i
                requiredCols.Add "Order ID", i
            
            Case "order date", "orderdate", "order_date", "date"
                orderDateCol = i
                requiredCols.Add "Order Date", i
                
            Case "customer", "customer id", "customerid", "customer_id"
                customerCol = i
                requiredCols.Add "Customer", i
                
            Case "product", "product id", "productid", "product_id"
                productCol = i
                requiredCols.Add "Product", i
                
            Case "quantity", "qty"
                quantityCol = i
                requiredCols.Add "Quantity", i
                
            Case "price", "unit price", "unitprice", "unit_price"
                priceCol = i
                requiredCols.Add "Price", i
        End Select
    Next i
    
    ' Check if all required columns exist
    Dim missingCols As String
    missingCols = ""
    
    If orderIdCol = 0 Then missingCols = missingCols & "Order ID, "
    If orderDateCol = 0 Then missingCols = missingCols & "Order Date, "
    If customerCol = 0 Then missingCols = missingCols & "Customer, "
    If productCol = 0 Then missingCols = missingCols & "Product, "
    If quantityCol = 0 Then missingCols = missingCols & "Quantity, "
    If priceCol = 0 Then missingCols = missingCols & "Price, "
    
    If Len(missingCols) > 0 Then
        missingCols = Left(missingCols, Len(missingCols) - 2) ' Remove trailing comma and space
        MsgBox "Required columns not found: " & missingCols, vbExclamation
        GoTo CleanExit
    End If
    
    ' Determine current quarter date range
    Dim currentYear As Integer, currentQuarter As Integer
    Dim startQuarterDate As Date, endQuarterDate As Date
    
    currentYear = Year(Date)
    currentQuarter = Int((Month(Date) - 1) / 3) + 1
    
    Select Case currentQuarter
        Case 1
            startQuarterDate = DateSerial(currentYear, 1, 1)
            endQuarterDate = DateSerial(currentYear, 3, 31)
        Case 2
            startQuarterDate = DateSerial(currentYear, 4, 1)
            endQuarterDate = DateSerial(currentYear, 6, 30)
        Case 3
            startQuarterDate = DateSerial(currentYear, 7, 1)
            endQuarterDate = DateSerial(currentYear, 9, 30)
        Case 4
            startQuarterDate = DateSerial(currentYear, 10, 1)
            endQuarterDate = DateSerial(currentYear, 12, 31)
    End Select
    
    ' Set up the validation results header
    With resultSheet
        .Range("A1").Value = "Order Data Validation Results"
        .Range("A3").Value = "Date:"
        .Range("B3").Value = Date
        .Range("C3").Value = "Time:"
        .Range("D3").Value = Time
        .Range("A5").Value = "Current Quarter:"
        .Range("B5").Value = "Q" & currentQuarter & " " & currentYear
        .Range("C5").Value = "Date Range:"
        .Range("D5").Value = Format(startQuarterDate, "yyyy-mm-dd") & " to " & Format(endQuarterDate, "yyyy-mm-dd")
        
        .Range("A7").Value = "Row"
        .Range("B7").Value = "Order ID"
        .Range("C7").Value = "Order Date"
        .Range("D7").Value = "Customer"
        .Range("E7").Value = "Product"
        .Range("F7").Value = "Issue Type"
        .Range("G7").Value = "Details"
        
        .Range("A1:G7").Font.Bold = True
    End With
    
    ' Store order
\`\`\`
`
