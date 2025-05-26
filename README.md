# exaMCP
## 📁 项目框架结构图

```
excel-automation-mcp/
├── 📁 backend/
│   ├── 📄 main.go                           # [1] 应用入口
│   ├── 📁 config/
│   │   ├── 📄 app_config.go                 # [2] 应用配置
│   │   └── 📄 logger_config.go              # [3] 日志配置
│   ├── 📁 models/
│   │   ├── 📄 excel_models.go               # [4] Excel数据模型
│   │   ├── 📄 mcp_models.go                 # [5] MCP相关模型
│   │   └── 📄 response_models.go            # [6] API响应模型
│   ├── 📁 service/
│   │   ├── 📁 excel/
│   │   │   ├── 📄 com_wrapper.go            # [7] Excel COM封装
│   │   │   ├── 📄 file_analyzer.go          # [8] Excel文件分析
│   │   │   ├── 📄 range_detector.go         # [9] 数据区域检测
│   │   │   ├── 📄 structure_parser.go       # [10] 结构解析
│   │   │   └── 📄 vba_executor.go           # [11] VBA执行器
│   │   ├── 📁 mcp/
│   │   │   ├── 📄 prompt_builder.go         # [12] 提示构建器
│   │   │   ├── 📄 context_manager.go        # [13] 上下文管理
│   │   │   └── 📄 tag_parser.go             # [14] 标签解析器
│   │   ├── 📁 llm/
│   │   │   ├── 📄 client_interface.go       # [15] LLM客户端接口
│   │   │   ├── 📄 openai_client.go          # [16] OpenAI客户端
│   │   │   └── 📄 claude_client.go          # [17] Claude客户端
│   │   ├── 📁 validation/
│   │   │   ├── 📄 vba_validator.go          # [18] VBA验证器
│   │   │   └── 📄 security_checker.go       # [19] 安全检查器
│   │   └── 📁 utils/
│   │       ├── 📄 error_handler.go          # [20] 错误处理器
│   │       └── 📄 file_utils.go             # [21] 文件工具
│   └── 📁 api/
│       ├── 📄 handlers.go                   # [22] API处理器
│       └── 📄 middleware.go                 # [23] 中间件
├── 📁 frontend/
│   ├── 📄 index.html                        # [24] 主页面
│   ├── 📄 debug.html                        # [25] 调试页面
│   ├── 📁 js/
│   │   ├── 📄 main.js                       # [26] 主逻辑
│   │   ├── 📄 excel-viewer.js               # [27] Excel查看器
│   │   ├── 📄 code-editor.js                # [28] 代码编辑器
│   │   └── 📄 api-client.js                 # [29] API客户端
│   ├── 📁 css/
│   │   ├── 📄 main.css                      # [30] 主样式
│   │   ├── 📄 components.css                # [31] 组件样式
│   │   └── 📄 prism.css                     # [32] 代码高亮样式
│   └── 📁 libs/
│       ├── 📄 prism.js                      # [33] 代码高亮库
│       └── 📄 xlsx.min.js                   # [34] Excel预览库
├── 📁 assets/
│   └── 📄 templates/                        # [35] VBA模板库
├── 📁 build/                                # 构建输出
├── 📄 wails.json                            # [36] Wails配置
├── 📄 go.mod                               # [37] Go模块配置
├── 📄 go.sum                               # Go依赖锁定
├── 📄 package.json                         # [38] 前端包配置
└── 📄 README.md                            # 项目说明
```
