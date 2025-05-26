package main

import (
	"context"
	"fmt"
	"log"
	"os"
	"runtime"
	"time"

	"github.com/wails-io/wails/v2"
	"github.com/wails-io/wails/v2/pkg/options"
	"github.com/wails-io/wails/v2/pkg/options/assetserver"
)

// App struct represents the main application
type App struct {
	ctx    context.Context
	cancel context.CancelFunc
	logger *log.Logger
}

// AppMetadata contains application information
type AppMetadata struct {
	Name        string `json:"name"`
	Version     string `json:"version"`
	Description string `json:"description"`
	Platform    string `json:"platform"`
	StartTime   string `json:"startTime"`
}

// NewApp creates a new App application struct
func NewApp() *App {
	logger := log.New(os.Stdout, "[ExcelMCP] ", log.LstdFlags|log.Lshortfile)
	
	return &App{
		logger: logger,
	}
}

// OnStartup is called when the app starts up. It sets up the application context
// and performs any initialization required
func (a *App) OnStartup(ctx context.Context) {
	// Create cancellable context for graceful shutdown
	a.ctx, a.cancel = context.WithCancel(ctx)
	
	a.logger.Println("Application starting up...")
	
	// Check if we're running on Windows (required for Excel COM)
	if runtime.GOOS != "windows" {
		a.logger.Println("WARNING: Excel COM functionality requires Windows platform")
	}
	
	// Initialize application components here
	a.initializeServices()
	
	a.logger.Println("Application startup completed successfully")
}

// OnShutdown is called when the app is about to quit,
// either by clicking the window close button or calling runtime.Quit
func (a *App) OnShutdown(ctx context.Context) {
	a.logger.Println("Application shutting down...")
	
	// Cancel any ongoing operations
	if a.cancel != nil {
		a.cancel()
	}
	
	// Cleanup resources here
	a.cleanupServices()
	
	a.logger.Println("Application shutdown completed")
}

// OnDomReady is called after front-end resources have been loaded
func (a *App) OnDomReady(ctx context.Context) {
	a.logger.Println("DOM Ready - Frontend loaded successfully")
}

// OnBeforeClose is called when the application is about to quit,
// either by clicking the window close button or calling runtime.Quit.
// Returning true will cause the application to continue, false will continue shutdown normally.
func (a *App) OnBeforeClose(ctx context.Context) (prevent bool) {
	a.logger.Println("Checking if application can close safely...")
	
	// Add any checks here to prevent closing if operations are in progress
	// For now, allow closing
	return false
}

// initializeServices initializes all application services
func (a *App) initializeServices() {
	a.logger.Println("Initializing application services...")
	
	// TODO: Initialize services in next development phase:
	// - Configuration service
	// - Excel service
	// - MCP service
	// - LLM service
	// - Validation service
	
	a.logger.Println("Services initialization completed")
}

// cleanupServices performs cleanup of all application services
func (a *App) cleanupServices() {
	a.logger.Println("Cleaning up application services...")
	
	// TODO: Cleanup services:
	// - Close Excel COM objects
	// - Cancel LLM requests
	// - Save application state
	// - Close database connections
	
	a.logger.Println("Services cleanup completed")
}

// API Methods for Frontend

// GetAppInfo returns basic application information
func (a *App) GetAppInfo() AppMetadata {
	return AppMetadata{
		Name:        "Excel Automation MCP",
		Version:     "1.0.0-dev",
		Description: "AI-powered Excel VBA code generator using Model Context Protocol",
		Platform:    runtime.GOOS,
		StartTime:   time.Now().Format("2006-01-02 15:04:05"),
	}
}

// HealthCheck returns application health status
func (a *App) HealthCheck() map[string]interface{} {
	return map[string]interface{}{
		"status":    "healthy",
		"timestamp": time.Now().Unix(),
		"platform":  runtime.GOOS,
		"goroutines": runtime.NumGoroutine(),
		"services": map[string]bool{
			"excel":      runtime.GOOS == "windows", // Excel COM only on Windows
			"config":     true,                      // Always available
			"mcp":        true,                      // Always available
			"validation": true,                      // Always available
		},
	}
}

// LogMessage logs a message from the frontend
func (a *App) LogMessage(level string, message string) {
	switch level {
	case "error":
		a.logger.Printf("FRONTEND ERROR: %s", message)
	case "warn":
		a.logger.Printf("FRONTEND WARN: %s", message)
	case "info":
		a.logger.Printf("FRONTEND INFO: %s", message)
	case "debug":
		a.logger.Printf("FRONTEND DEBUG: %s", message)
	default:
		a.logger.Printf("FRONTEND: %s", message)
	}
}

// GetSystemInfo returns detailed system information
func (a *App) GetSystemInfo() map[string]interface{} {
	return map[string]interface{}{
		"os":           runtime.GOOS,
		"arch":         runtime.GOARCH,
		"goVersion":    runtime.Version(),
		"cpus":         runtime.NumCPU(),
		"goroutines":   runtime.NumGoroutine(),
		"memStats": func() map[string]interface{} {
			var m runtime.MemStats
			runtime.ReadMemStats(&m)
			return map[string]interface{}{
				"alloc":      m.Alloc,
				"totalAlloc": m.TotalAlloc,
				"sys":        m.Sys,
				"numGC":      m.NumGC,
			}
		}(),
	}
}

// Error handling and recovery
func (a *App) handlePanic() {
	if r := recover(); r != nil {
		a.logger.Printf("PANIC RECOVERED: %v", r)
		// Add additional panic handling here
		// Could send crash reports, save state, etc.
	}
}

// Graceful error handling wrapper
func (a *App) safeExecute(operation string, fn func() error) error {
	defer a.handlePanic()
	
	a.logger.Printf("Executing operation: %s", operation)
	
	if err := fn(); err != nil {
		a.logger.Printf("Operation %s failed: %v", operation, err)
		return fmt.Errorf("operation %s failed: %w", operation, err)
	}
	
	a.logger.Printf("Operation %s completed successfully", operation)
	return nil
}

// main function - application entry point
func main() {
	// Create an instance of the app structure
	app := NewApp()

	// Create application with options
	err := wails.Run(&options.App{
		Title:  "Excel Automation MCP",
		Width:  1200,
		Height: 800,
		MinWidth: 800,
		MinHeight: 600,
		AssetServer: &assetserver.Options{
			Assets: nil, // Will be set by Wails build process
		},
		BackgroundColour: &options.RGBA{R: 27, G: 38, B: 54, A: 1},
		OnStartup:        app.OnStartup,
		OnDomReady:       app.OnDomReady,
		OnBeforeClose:    app.OnBeforeClose,
		OnShutdown:       app.OnShutdown,
		WindowStartState: options.Normal,
		Frameless:        false,
		CSSDragProperty:  "widows",
		CSSDragValue:     "1",
		Debug: options.Debug{
			OpenInspectorOnStartup: false,
		},
		// Export app methods to frontend
		Ctx: app,
	})

	if err != nil {
		log.Fatal("Error starting application:", err)
	}
}
