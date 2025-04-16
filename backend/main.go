package main

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"runtime/debug"
	"sync"
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/rs/zerolog"
	"github.com/rs/zerolog/log"
	"github.com/spf13/viper"
	"github.com/wailsapp/wails/v2"
	"github.com/wailsapp/wails/v2/pkg/options"
	"github.com/wailsapp/wails/v2/pkg/options/assetserver"
	"github.com/wailsapp/wails/v2/pkg/runtime"

	"github.com/exaMCP/backend/service/excel"
	"github.com/exaMCP/backend/service/feedback"
	"github.com/exaMCP/backend/service/llm"
	"github.com/exaMCP/backend/service/mcp"
	"github.com/exaMCP/backend/service/validator"
)

// Config 应用配置结构
type Config struct {
	LLM struct {
		Provider    string  `mapstructure:"provider"`
		APIKey      string  `mapstructure:"api_key"`
		Temperature float64 `mapstructure:"temperature"`
		MaxTokens   int     `mapstructure:"max_tokens"`
		Endpoint    string  `mapstructure:"endpoint"`
	} `mapstructure:"llm"`

	Excel struct {
		MaxSheets   int `mapstructure:"max_sheets"`
		MaxRows     int `mapstructure:"max_rows"`
		MaxSamples  int `mapstructure:"max_samples"`
		UseAdvanced bool `mapstructure:"use_advanced_detection"`
	} `mapstructure:"excel"`

	UI struct {
		Theme             string `mapstructure:"theme"`
		CodeHighlighting  bool   `mapstructure:"code_highlighting"`
		ShowLineNumbers   bool   `mapstructure:"show_line_numbers"`
		DefaultLanguage   string `mapstructure:"default_language"`
		AutosaveInterval  int    `mapstructure:"autosave_interval"`
	} `mapstructure:"ui"`

	Logging struct {
		Level        string `mapstructure:"level"`
		File         string `mapstructure:"file"`
		EnableConsole bool   `mapstructure:"enable_console"`
		MaxSize      int    `mapstructure:"max_size"`
		MaxBackups   int    `mapstructure:"max_backups"`
		MaxAge       int    `mapstructure:"max_age"`
	} `mapstructure:"logging"`

	System struct {
		DevMode        bool   `mapstructure:"dev_mode"`
		TempDir        string `mapstructure:"temp_dir"`
		MaxConcurrency int    `mapstructure:"max_concurrency"`
		CacheEnabled   bool   `mapstructure:"cache_enabled"`
		CacheDir       string `mapstructure:"cache_dir"`
	} `mapstructure:"system"`
}

// App 表示应用程序主结构
type App struct {
	ctx             context.Context
	config          *Config
	excelAnalysis   *excel.ExcelStructureAnalysis
	currentRegion   int
	filePath        string
	llmClient       *llm.Client
	mutex           sync.Mutex
	taskQueue       chan Task
	workingDir      string
	analyticsLogger *zerolog.Logger
	sessionStarted  time.Time
	sessionID       string
}

// Task 表示一个后台任务
type Task struct {
	ID          string
	Type        string
	Description string
	Status      string
	StartTime   time.Time
	EndTime     time.Time
	Error       error
}

// ErrorResponse 标准错误响应结构
type ErrorResponse struct {
	Code    string `json:"code"`
	Message string `json:"message"`
	Details any    `json:"details,omitempty"`
}

// NewApp 创建应用程序实例
func NewApp() *App {
	// 初始化默认工作目录
	homeDir, err := os.UserHomeDir()
	if err != nil {
		panic(fmt.Errorf("获取用户主目录失败: %w", err))
	}

	workingDir := filepath.Join(homeDir, ".examcp")
	if err := os.MkdirAll(workingDir, 0755); err != nil {
		panic(fmt.Errorf("创建工作目录失败: %w", err))
	}

	// 初始化应用程序实例
	app := &App{
		currentRegion:  -1,
		taskQueue:      make(chan Task, 100),
		workingDir:     workingDir,
		sessionStarted: time.Now(),
		sessionID:      fmt.Sprintf("session-%d", time.Now().Unix()),
	}

	// 加载配置
	app.loadConfig()

	// 初始化日志
	app.initLogging()

	log.Info().
		Str("sessionID", app.sessionID).
		Msg("应用程序初始化完成")

	// 启动后台任务处理器
	go app.taskProcessor()

	return app
}

// loadConfig 加载应用程序配置
func (a *App) loadConfig() {
	// 设置默认配置
	viper.SetDefault("llm.provider", "openrouter")
	viper.SetDefault("llm.temperature", 0.7)
	viper.SetDefault("llm.max_tokens", 4000)
	viper.SetDefault("excel.max_sheets", 20)
	viper.SetDefault("excel.max_rows", 5000)
	viper.SetDefault("excel.max_samples", 5)
	viper.SetDefault("excel.use_advanced_detection", true)
	viper.SetDefault("ui.theme", "light")
	viper.SetDefault("ui.code_highlighting", true)
	viper.SetDefault("ui.show_line_numbers", true)
	viper.SetDefault("ui.default_language", "zh_CN")
	viper.SetDefault("ui.autosave_interval", 300)
	viper.SetDefault("logging.level", "info")
	viper.SetDefault("logging.file", filepath.Join(a.workingDir, "logs/app.log"))
	viper.SetDefault("logging.enable_console", true)
	viper.SetDefault("logging.max_size", 10)
	viper.SetDefault("logging.max_backups", 3)
	viper.SetDefault("logging.max_age", 30)
	viper.SetDefault("system.dev_mode", false)
	viper.SetDefault("system.temp_dir", filepath.Join(a.workingDir, "temp"))
	viper.SetDefault("system.max_concurrency", 4)
	viper.SetDefault("system.cache_enabled", true)
	viper.SetDefault("system.cache_dir", filepath.Join(a.workingDir, "cache"))

	// 设置配置名
	viper.SetConfigName("config")
	viper.SetConfigType("yaml")

	// 按优先级添加配置搜索路径
	viper.AddConfigPath(a.workingDir)
	if configDir, err := os.UserConfigDir(); err == nil {
		viper.AddConfigPath(filepath.Join(configDir, "examcp"))
	}

	// 读取配置
	if err := viper.ReadInConfig(); err != nil {
		// 忽略配置文件不存在错误，会使用默认值
		if _, ok := err.(viper.ConfigFileNotFoundError); !ok {
			fmt.Printf("配置文件读取错误: %v\n", err)
		}

		// 写入默认配置
		if err := os.MkdirAll(a.workingDir, 0755); err == nil {
			viper.SafeWriteConfig()
		}
	}

	// 加载配置到结构
	a.config = &Config{}
	if err := viper.Unmarshal(a.config); err != nil {
		fmt.Printf("配置解析错误: %v\n", err)
	}

	// 创建子目录
	dirs := []string{
		a.config.System.TempDir,
		a.config.System.CacheDir,
		filepath.Dir(a.config.Logging.File),
	}

	for _, dir := range dirs {
		if dir != "" {
			if err := os.MkdirAll(dir, 0755); err != nil {
				fmt.Printf("创建目录失败 %s: %v\n", dir, err)
			}
		}
	}
}

// initLogging 初始化日志系统
func (a *App) initLogging() {
	// 设置日志格式
	zerolog.TimeFieldFormat = time.RFC3339

	// 确定日志级别
	level, err := zerolog.ParseLevel(a.config.Logging.Level)
	if err != nil {
		level = zerolog.InfoLevel
	}
	zerolog.SetGlobalLevel(level)

	// 创建日志输出
	var outputs []zerolog.Logger
	
	// 控制台输出
	if a.config.Logging.EnableConsole {
		consoleWriter := zerolog.ConsoleWriter{Out: os.Stdout, TimeFormat: time.RFC3339}
		outputs = append(outputs, zerolog.New(consoleWriter).With().Timestamp().Logger())
	}

	// 文件输出
	if a.config.Logging.File != "" {
		logDir := filepath.Dir(a.config.Logging.File)
		if err := os.MkdirAll(logDir, 0755); err != nil {
			fmt.Printf("创建日志目录失败: %v\n", err)
		} else {
			logFile, err := os.OpenFile(a.config.Logging.File, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0666)
			if err != nil {
				fmt.Printf("打开日志文件失败: %v\n", err)
			} else {
				outputs = append(outputs, zerolog.New(logFile).With().Timestamp().Logger())
			}
		}
	}

	// 创建多输出
	if len(outputs) > 0 {
		multi := zerolog.MultiLevelWriter(consoleOutput(outputs))
		log.Logger = zerolog.New(multi).With().
			Str("service", "examcp").
			Str("session", a.sessionID).
			Timestamp().
			Logger()
	}

	// 单独的分析日志
	analyticsLogFile := filepath.Join(filepath.Dir(a.config.Logging.File), "analytics.log")
	if analyticsFile, err := os.OpenFile(analyticsLogFile, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0666); err == nil {
		analyticsLogger := zerolog.New(analyticsFile).With().
			Str("type", "analytics").
			Str("session", a.sessionID).
			Timestamp().
			Logger()
		a.analyticsLogger = &analyticsLogger
	}
}

// consoleOutput 将多个日志转换为多级别写入器
func consoleOutput(loggers []zerolog.Logger) []zerolog.LevelWriter {
	writers := make([]zerolog.LevelWriter, len(loggers))
	for i, logger := range loggers {
		writers[i] = logger
	}
	return writers
}

// startup 应用程序启动时调用
func (a *App) startup(ctx context.Context) {
	a.ctx = ctx
	
	// 初始化LLM客户端
	a.llmClient = llm.NewClient(a.config.LLM.Provider, a.config.LLM.APIKey, a.config.LLM.Endpoint)
	a.llmClient.SetTemperature(a.config.LLM.Temperature)
	a.llmClient.SetMaxTokens(a.config.LLM.MaxTokens)
	
	// 记录启动事件
	log.Info().
		Str("configDir", a.workingDir).
		Str("provider", a.config.LLM.Provider).
		Str("theme", a.config.UI.Theme).
		Bool("devMode", a.config.System.DevMode).
		Msg("应用程序启动")
	
	// 显示欢迎消息
	runtime.EventsEmit(a.ctx, "app:notification", map[string]interface{}{
		"type":    "info",
		"title":   "准备就绪",
		"message": "exaMCP 已启动",
	})
}

// shutdown 应用程序关闭时调用
func (a *App) shutdown(ctx context.Context) {
	// 记录会话时长
	sessionDuration := time.Since(a.sessionStarted).Seconds()
	log.Info().
		Float64("durationSeconds", sessionDuration).
		Msg("应用程序关闭")
	
	// 记录分析数据
	if a.analyticsLogger != nil {
		(*a.analyticsLogger).Info().
			Float64("sessionDuration", sessionDuration).
			Str("sessionID", a.sessionID).
			Msg("session_end")
	}
	
	// 清理临时文件
	if tempDir := a.config.System.TempDir; tempDir != "" {
		os.RemoveAll(tempDir)
	}
}

// taskProcessor 后台任务处理器
func (a *App) taskProcessor() {
	for task := range a.taskQueue {
		log.Debug().
			Str("taskID", task.ID).
			Str("taskType", task.Type).
			Msg("开始处理任务")
		
		// 更新任务状态
		task.Status = "processing"
		runtime.EventsEmit(a.ctx, "task:update", task)
		
		// 执行任务
		func() {
			defer func() {
				if r := recover(); r != nil {
					log.Error().
						Str("taskID", task.ID).
						Interface("panic", r).
						Str("stack", string(debug.Stack())).
						Msg("任务执行发生崩溃")
					
					task.Status = "failed"
					task.Error = fmt.Errorf("内部错误: %v", r)
					runtime.EventsEmit(a.ctx, "task:update", task)
				}
			}()
			
			// 任务执行逻辑 (实际应用中根据task.Type进行分发)
			time.Sleep(100 * time.Millisecond) // 模拟任务执行
		}()
		
		// 完成任务
		if task.Error == nil {
			task.Status = "completed"
		} else {
			task.Status = "failed"
		}
		
		task.EndTime = time.Now()
		runtime.EventsEmit(a.ctx, "task:update", task)
		
		log.Debug().
			Str("taskID", task.ID).
			Str("taskType", task.Type).
			Str("status", task.Status).
			Dur("duration", task.EndTime.Sub(task.StartTime)).
			Msg("任务处理完成")
	}
}

// SelectExcelFile 打开文件选择对话框选择Excel文件
func (a *App) SelectExcelFile() (string, error) {
	defer a.recoverPanic("SelectExcelFile")
	
	a.mutex.Lock()
	defer a.mutex.Unlock()
	
	log.Debug().Msg("打开文件选择对话框")
	
	filters := []runtime.FileFilter{
		{
			DisplayName: "Excel Files (*.xlsx;*.xls)",
			Pattern:     "*.xlsx;*.xls",
		},
		{
			DisplayName: "All Files (*.*)",
			Pattern:     "*.*",
		},
	}
	
	filePath, err := runtime.OpenFileDialog(a.ctx, runtime.OpenDialogOptions{
		Title:   "选择Excel文件",
		Filters: filters,
	})
	
	if err != nil {
		log.Error().Err(err).Msg("文件选择失败")
		return "", fmt.Errorf("选择文件时出错: %w", err)
	}
	
	if filePath == "" {
		log.Debug().Msg("用户取消了文件选择")
		return "", nil
	}
	
	log.Info().Str("filePath", filePath).Msg("用户选择了Excel文件")
	a.filePath = filePath
	
	return filePath, nil
}

// AnalyzeExcelFile 分析Excel文件结构
func (a *App) AnalyzeExcelFile(filePath string) (*excel.ExcelStructureAnalysis, error) {
	defer a.recoverPanic("AnalyzeExcelFile")
	
	log.Info().Str("filePath", filePath).Msg("开始分析Excel文件")
	
	// 创建并发送任务
	task := Task{
		ID:          fmt.Sprintf("analyze-%d", time.Now().UnixNano()),
		Type:        "excel_analysis",
		Description: "分析Excel文件结构",
		Status:      "pending",
		StartTime:   time.Now(),
	}
	
	// 通知前端任务开始
	runtime.EventsEmit(a.ctx, "task:start", task)
	
	// 配置分析选项
	options := excel.AnalysisOptions{
		MaxSheets:   a.config.Excel.MaxSheets,
		MaxRows:     a.config.Excel.MaxRows,
		MaxSamples:  a.config.Excel.MaxSamples,
		UseAdvanced: a.config.Excel.UseAdvanced,
	}
	
	// 使用缓存
	var analysis *excel.ExcelStructureAnalysis
	var err error
	
	if a.config.System.CacheEnabled {
		cacheFile := filepath.Join(a.config.System.CacheDir, fmt.Sprintf("%x.json", excel.HashFilePath(filePath)))
		analysis, err = excel.LoadAnalysisFromCache(cacheFile)
		
		if err == nil {
			log.Debug().Str("cacheFile", cacheFile).Msg("从缓存加载Excel分析结果")
		} else {
			// 缓存未命中，执行分析
			analysis, err = excel.AnalyzeExcelFile(filePath, options)
			if err == nil && analysis != nil {
				// 保存到缓存
				excel.SaveAnalysisToCache(analysis, cacheFile)
			}
		}
	} else {
		// 禁用缓存，直接分析
		analysis, err = excel.AnalyzeExcelFile(filePath, options)
	}
	
	// 更新任务状态
	task.EndTime = time.Now()
	if err != nil {
		task.Status = "failed"
		task.Error = err
		log.Error().Err(err).Str("filePath", filePath).Msg("Excel分析失败")
	} else {
		task.Status = "completed"
		log.Info().
			Int("sheets", len(analysis.SheetNames)).
			Int("ranges", len(analysis.AllRanges)).
			Str("filePath", filePath).
			Msg("Excel分析完成")
	}
	
	// 通知前端任务完成
	runtime.EventsEmit(a.ctx, "task:update", task)
	
	// 记录分析结果
	if err == nil && analysis != nil {
		a.excelAnalysis = analysis
		a.filePath = filePath
		
		// 记录分析指标
		if a.analyticsLogger != nil {
			(*a.analyticsLogger).Info().
				Str("action", "excel_analysis").
				Int("sheetCount", len(analysis.SheetNames)).
				Int("rangeCount", len(analysis.AllRanges)).
				Int("totalHeaders", a.countTotalHeaders(analysis)).
				Str("filePath", filepath.Base(filePath)).
				Msg("excel_analyzed")
		}
	}
	
	if err != nil {
		return nil, a.wrapError("EXCEL_ANALYSIS_FAILED", err, nil)
	}
	
	return analysis, nil
}

// countTotalHeaders 计算所有区域的表头总数
func (a *App) countTotalHeaders(analysis *excel.ExcelStructureAnalysis) int {
	totalHeaders := 0
	for _, rng := range analysis.AllRanges {
		totalHeaders += len(rng.Headers)
	}
	return totalHeaders
}

// SetCurrentRegion 设置当前选中的区域
func (a *App) SetCurrentRegion(index int) bool {
	defer a.recoverPanic("SetCurrentRegion")
	
	log.Debug().Int("index", index).Msg("设置当前区域")
	
	if a.excelAnalysis == nil || index < 0 || index >= len(a.excelAnalysis.AllRanges) {
		log.Warn().Int("index", index).Msg("无效的区域索引")
		return false
	}
	
	a.currentRegion = index
	
	// 记录区域选择
	region := a.excelAnalysis.AllRanges[index]
	log.Info().
		Int("index", index).
		Str("range", region.RangeAddress).
		Str("sheet", region.SheetName).
		Int("headers", len(region.Headers)).
		Msg("用户选择了数据区域")
	
	return true
}

// PreviewRegion 获取区域预览数据
func (a *App) PreviewRegion(index int) ([][]string, error) {
	defer a.recoverPanic("PreviewRegion")
	
	log.Debug().Int("index", index).Msg("获取区域预览")
	
	if a.excelAnalysis == nil || index < 0 || index >= len(a.excelAnalysis.AllRanges) {
		return nil, a.wrapError("INVALID_REGION", fmt.Errorf("无效的区域索引"), map[string]interface{}{
			"index":        index,
			"regionCount":  len(a.excelAnalysis.AllRanges),
			"hasAnalysis": a.excelAnalysis != nil,
		})
	}
	
	// 获取指定区域
	region := a.excelAnalysis.AllRanges[index]
	
	// 从Excel文件中读取预览数据
	preview, err := excel.GetRegionPreviewData(a.filePath, region.SheetName, region.RangeAddress, region.HasHeaders, 10)
	if err != nil {
		log.Error().
			Err(err).
			Int("index", index).
			Str("range", region.RangeAddress).
			Str("sheet", region.SheetName).
			Msg("获取预览数据失败")
		return nil, a.wrapError("PREVIEW_FAILED", err, nil) 
	}
	
	return preview, nil
}

// GenerateVBAScript 生成VBA脚本
func (a *App) GenerateVBAScript(description string) (string, error) {
	defer a.recoverPanic("GenerateVBAScript")
	
	log.Info().
		Str("description", description).
		Int("length", len(description)).
		Msg("开始生成VBA脚本")
	
	if a.excelAnalysis == nil {
		return "", a.wrapError("NO_EXCEL_DATA", fmt.Errorf("请先导入Excel文件"), nil)
	}
	
	if a.currentRegion < 0 || a.currentRegion >= len(a.excelAnalysis.AllRanges) {
		return "", a.wrapError("NO_REGION_SELECTED", fmt.Errorf("请先选择数据区域"), nil)
	}
	
	// 获取当前区域
	region := a.excelAnalysis.AllRanges[a.currentRegion]
	
	// 创建任务
	task := Task{
		ID:          fmt.Sprintf("generate-%d", time.Now().UnixNano()),
		Type:        "vba_generation",
		Description: "生成VBA脚本",
		Status:      "pending",
		StartTime:   time.Now(),
	}
	
	// 通知前端任务开始
	runtime.EventsEmit(a.ctx, "task:start", task)
	
	// 创建适用于此区域的MCP提示
	var prompt string
	if a.config.Excel.UseAdvanced {
		prompt = mcp.AdvancedMCPPrompt(region, description, true)
	} else {
		prompt = mcp.GenerateMCPPrompt(region, description, true)
	}
	
	// 记录提示长度
	log.Debug().
		Int("promptLength", len(prompt)).
		Str("region", region.RangeAddress).
		Msg("生成提示完成")
	
	// 调用LLM API生成VBA代码
	vbaCode, err := a.llmClient.Generate(prompt)
	if err != nil {
		log.Error().Err(err).Msg("LLM API调用失败")
		
		task.Status = "failed"
		task.Error = err
		task.EndTime = time.Now()
		runtime.EventsEmit(a.ctx, "task:update", task)
		
		return "", a.wrapError("LLM_API_ERROR", err, nil)
	}
	
	// 验证生成的VBA代码
	validationResult := validator.ValidateVBACode(vbaCode)
	
	if !validationResult.IsValid {
		log.Warn().
			Interface("warnings", validationResult.Warnings).
			Int("safetyScore", validationResult.SafetyScore).
			Msg("VBA代码验证存在问题")
		
		// 尝试修复代码
		improvedCode, fixErr := validator.SuggestCodeImprovements(vbaCode, validationResult)
		if fixErr == nil {
			log.Info().Msg("VBA代码已自动改进")
			vbaCode = improvedCode
		} else {
			log.Warn().Err(fixErr).Msg("代码自动改进失败")
		}
		
		// 提示用户检查代码
		runtime.EventsEmit(a.ctx, "app:notification", map[string]interface{}{
			"type":    "warning",
			"title":   "请检查生成的代码",
			"message": "代码验证发现潜在问题，请仔细检查。",
			"details": validationResult.Warnings,
		})
	}
	
	// 记录生成度量
	taskDuration := time.Since(task.StartTime)
	
	// 更新任务状态
	task.Status = "completed"
	task.EndTime = time.Now()
	runtime.EventsEmit(a.ctx, "task:update", task)
	
	// 记录分析事件
	if a.analyticsLogger != nil {
		(*a.analyticsLogger).Info().
			Str("action", "vba_generation").
			Int("promptLength", len(prompt)).
			Int("codeLength", len(vbaCode)).
			Float64("durationSec", taskDuration.Seconds()).
			Int("safetyScore", validationResult.SafetyScore).
			Msg("vba_generated")
	}
	
	log.Info().
		Int("codeLength", len(vbaCode)).
		Float64("durationSec", taskDuration.Seconds()).
		Int("warnings", len(validationResult.Warnings)).
		Int("suggestions", len(validationResult.Suggestions)).
		Msg("VBA代码生成完成")
	
	return vbaCode, nil
}

// ExecuteVBA 执行VBA代码
func (a *App) ExecuteVBA(vbaCode string) (string, error) {
	defer a.recoverPanic("ExecuteVBA")
	
	if a.filePath == "" {
		return "", a.wrapError("NO_FILE", fmt.Errorf("未选择Excel文件"), nil)
	}
	
	log.Info().
		Int("codeLength", len(vbaCode)).
		Str("filePath", a.filePath).
		Msg("开始执行VBA代码")
	
	// 创建任务
	task := Task{
		ID:          fmt.Sprintf("execute-%d", time.Now().UnixNano()),
		Type:        "vba_execution",
		Description: "执行VBA脚本",
		Status:      "pending",
		StartTime:   time.Now(),
	}
	
	// 通知前端任务开始
	runtime.EventsEmit(a.ctx, "task:start", task)
	
	// 尝试执行VBA代码并捕获潜在错误
	errorInfo, err := excel.ExecuteVBAWithErrorAnalysis(a.filePath, vbaCode)
	
	// 更新任务状态
	task.EndTime = time.Now()
	if err != nil {
		task.Status = "failed"
		task.Error = err
		runtime.EventsEmit(a.ctx, "task:update", task)
		
		// 如果有详细错误信息，返回给前端
		if errorInfo != nil {
			log.Error().
				Err(err).
				Str("errorMessage", errorInfo.ErrorMessage).
				Int("errorNumber", errorInfo.ErrorNumber).
				Str("procedure", errorInfo.ProcedureName).
				Int("lineNumber", errorInfo.LineNumber).
				Msg("VBA执行失败")
			
			return "", a.wrapError("VBA_EXECUTION_ERROR", err, errorInfo)
		}
		
		log.Error().Err(err).Msg("VBA执行失败但无详细错误信息")
		return "", a.wrapError("VBA_EXECUTION_ERROR", err, nil)
	}
	
	task.Status = "completed"
	runtime.EventsEmit(a.ctx, "task:update", task)
	
	// 记录执行事件
	if a.analyticsLogger != nil {
		(*a.analyticsLogger).Info().
			Str("action", "vba_execution").
			Bool("success", err == nil).
			Float64("durationSec", time.Since(task.StartTime).Seconds()).
			Msg("vba_executed")
	}
	
	log.Info().
		Str("filePath", a.filePath).
		Float64("durationSec", time.Since(task.StartTime).Seconds()).
		Msg("VBA代码执行成功")
	
	return "VBA代码执行成功", nil
}

// GetAutoFixSuggestion 获取代码自动修复建议
func (a *App) GetAutoFixSuggestion(vbaCode string, errorInfoJSON string) (string, error) {
	defer a.recoverPanic("GetAutoFixSuggestion")
	
	log.Info().
		Int("codeLength", len(vbaCode)).
		Int("errorInfoLength", len(errorInfoJSON)).
		Msg("请求代码修复建议")
	
	// 解析错误信息
	var errorInfo excel.VBAErrorInfo
	if err := json.Unmarshal([]byte(errorInfoJSON), &errorInfo); err != nil {
		return "", a.wrapError("INVALID_ERROR_INFO", err, nil)
	}
	
	// 创建反馈对象
	userFeedback := fmt.Sprintf("Fix this code issue: %s, at line %d, procedure: %s",
		errorInfo.ErrorMessage,
		errorInfo.LineNumber,
		errorInfo.ProcedureName)
	
	// 尝试获取修复建议
	refinedCode, err := feedback.RefineVBACode(vbaCode, userFeedback)
	if err != nil {
		log.Error().Err(err).Msg("获取代码修复建议失败")
		return "", a.wrapError("FIX_SUGGESTION_FAILED", err, nil)
	}
	
	// 记录修复事件
	log.Info().
		Str("errorMessage", errorInfo.ErrorMessage).
		Int("lineNumber", errorInfo.LineNumber).
		Bool("success", true).
		Msg("代码修复建议生成成功")
	
	// 记录分析事件
	if a.analyticsLogger != nil {
		(*a.analyticsLogger).Info().
			Str("action", "code_fix").
			Str("errorType", fmt.Sprintf("%d", errorInfo.ErrorNumber)).
			Bool("success", true).
			Msg("code_fix_generated")
	}
	
	return refinedCode, nil
}

// GetAppConfig 获取应用程序配置
func (a *App) GetAppConfig() map[string]interface{} {
	defer a.recoverPanic("GetAppConfig")
	
	// 返回不含敏感信息的配置
	return map[string]interface{}{
		"ui": a.config.UI,
		"system": map[string]interface{}{
			"devMode":      a.config.System.DevMode,
			"cacheEnabled": a.config.System.CacheEnabled,
		},
		"excel": map[string]interface{}{
			"useAdvanced": a.config.Excel.UseAdvanced,
		},
		"llm": map[string]interface{}{
			"provider": a.config.LLM.Provider,
		},
	}
}

// UpdateAppConfig 更新应用程序配置
func (a *App) UpdateAppConfig(configJSON string) error {
	defer a.recoverPanic("UpdateAppConfig")
	
	log.Debug().
		Int("jsonLength", len(configJSON)).
		Msg("更新应用程序配置")
	
	var configUpdate map[string]interface{}
	if err := json.Unmarshal([]byte(configJSON), &configUpdate); err != nil {
		return a.wrapError("INVALID_CONFIG", err, nil)
	}
	
	// 更新Viper配置
	for section, values := range configUpdate {
		if valuesMap, ok := values.(map[string]interface{}); ok {
			for key, value := range valuesMap {
				configKey := fmt.Sprintf("%s.%s", section, key)
				viper.Set(configKey, value)
				
				log.Debug().
					Str("key", configKey).
					Interface("value", value).
					Msg("更新配置项")
			}
		}
	}
	
	// 保存配置
	if err := viper.WriteConfig(); err != nil {
		log.Error().Err(err).Msg("保存配置失败")
		return a.wrapError("CONFIG_SAVE_FAILED", err, nil)
	}
	
	// 重新加载配置
	a.loadConfig()
	
	log.Info().Msg("应用程序配置已更新")
	return nil
}

// TestSQLQuery 测试SQL查询并返回预览结果
func (a *App) TestSQLQuery(filePath string, sqlQuery string) (string, error) {
	defer a.recoverPanic("TestSQLQuery")
	
	log.Info().
		Str("filePath", filePath).
		Str("sqlQuery", sqlQuery).
		Msg("测试SQL查询")
	
	// 验证文件路径
	if filePath == "" {
		return "", a.wrapError("NO_FILE", fmt.Errorf("未指定Excel文件"), nil)
	}
	
	// 验证SQL查询
	if sqlQuery == "" {
		return "", a.wrapError("EMPTY_QUERY", fmt.Errorf("SQL查询为空"), nil)
	}
	
	// 执行SQL查询测试
	result, err := excel.TestSQLQuery(filePath, sqlQuery)
	if err != nil {
		log.Error().
			Err(err).
			Str("sqlQuery", sqlQuery).
			Msg("SQL查询测试失败")
		return "", a.wrapError("SQL_QUERY_FAILED", err, nil)
	}
	
	log.Debug().
		Str("sqlQuery", sqlQuery).
		Int("resultLength", len(result)).
		Msg("SQL查询测试成功")
	
	return result, nil
}

// recoverPanic 恢复并记录崩溃
func (a *App) recoverPanic(methodName string) {
	if r := recover(); r != nil {
		log.Error().
			Str("method", methodName).
			Interface("panic", r).
			Str("stack", string(debug.Stack())).
			Msg("方法执行发生崩溃")
		
		// 通知前端
		if a.ctx != nil {
			runtime.EventsEmit(a.ctx, "app:error", map[string]interface{}{
				"code":    "INTERNAL_ERROR",
				"message": fmt.Sprintf("内部错误: %v", r),
			})
		}
	}
}

// wrapError 包装错误为标准格式
func (a *App) wrapError(code string, err error, details interface{}) error {
	errResp := ErrorResponse{
		Code:    code,
		Message: err.Error(),
		Details: details,
	}
	
	// 序列化为JSON
	jsonBytes, jsonErr := json.Marshal(errResp)
	if jsonErr != nil {
		return fmt.Errorf("错误(%s): %v", code, err)
	}
	
	return fmt.Errorf("%s", string(jsonBytes))
}

// main 应用程序入口
func main() {
	// 创建应用
	app := NewApp()
	
	// 运行应用
	err := wails.Run(&options.App{
		Title:             "exaMCP - Excel Automation",
		Width:             1024,
		Height:            768,
		MinWidth:          800,
		MinHeight:         600,
		MaxWidth:          0,
		MaxHeight:         0,
		DisableResize:     false,
		Fullscreen:        false,
		Frameless:         false,
		StartHidden:       false,
		HideWindowOnClose: false,
		BackgroundColour:  &options.RGBA{R: 255, G: 255, B: 255, A: 1},
		AssetServer:       &assetserver.Options{
			Assets:     os.DirFS("frontend/dist"),
			Handler:    nil,
			Middleware: nil,
		},
		Menu:             nil,
		Logger:           nil,
		LogLevel:         0,
		OnStartup:        app.startup,
		OnDomReady:       nil,
		OnShutdown:       app.shutdown,
		OnBeforeClose:    nil,
		EnableDefaultContextMenu: false,
		SingleInstance:   true,
		Bind: []interface{}{
			app,
		},
		Windows: &windows.Options{
			WebviewIsTransparent:              false,
			WindowIsTranslucent:               false,
			DisableWindowIcon:                 false,
			DisableFramelessWindowDecorations: false,
			WebviewUserDataPath:               "",
			Theme:                             windows.SystemDefault,
		},
		Mac: &mac.Options{
			TitleBar: &mac.TitleBar{
				TitleBarStyle:            mac.DefaultTitleBar,
				TitleBarHideInFullscreen: false,
				HideTitle:                false,
				HideTitleBar:             false,
				FullSizeContent:          false,
				UseToolbar:               false,
				HideToolbarSeparator:     false,
			},
			Appearance:           mac.DefaultAppearance,
			WebviewIsTransparent: false,
			WindowIsTranslucent:  false,
			About: &mac.AboutInfo{
				Title:   "exaMCP - Excel Automation with MCP",
				Message: "© 2025 Your Company",
				Icon:    nil,
			},
		},
		Linux: &linux.Options{
			Icon:                linux.MissingImageIcon,
			WindowIsTranslucent: false,
		},
	})
	
	if err != nil {
		log.Fatal().Err(err).Msg("应用程序启动失败")
	}
}
