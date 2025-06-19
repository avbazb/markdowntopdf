# Markdown转PDF工具构建脚本
# 用于构建和运行项目

param(
    [string]$Action = "build",  # build, run, clean, restore
    [string]$Configuration = "Debug"
)

Write-Host "Markdown转PDF工具构建脚本" -ForegroundColor Green
Write-Host "=========================" -ForegroundColor Green

# 检查.NET SDK
Write-Host "检查.NET SDK..." -ForegroundColor Yellow
try {
    $dotnetVersion = dotnet --version
    Write-Host ".NET SDK版本: $dotnetVersion" -ForegroundColor Green
}
catch {
    Write-Host "错误: 未找到.NET SDK，请先安装.NET 6.0 SDK" -ForegroundColor Red
    exit 1
}

# 项目文件路径
$projectFile = "MarkdownToPdf.csproj"

switch ($Action.ToLower()) {
    "restore" {
        Write-Host "正在还原NuGet包..." -ForegroundColor Yellow
        dotnet restore $projectFile
        if ($LASTEXITCODE -eq 0) {
            Write-Host "包还原完成!" -ForegroundColor Green
        } else {
            Write-Host "包还原失败!" -ForegroundColor Red
            exit 1
        }
    }
    
    "clean" {
        Write-Host "正在清理项目..." -ForegroundColor Yellow
        dotnet clean $projectFile --configuration $Configuration
        if ($LASTEXITCODE -eq 0) {
            Write-Host "清理完成!" -ForegroundColor Green
        } else {
            Write-Host "清理失败!" -ForegroundColor Red
            exit 1
        }
    }
    
    "build" {
        Write-Host "正在构建项目 ($Configuration)..." -ForegroundColor Yellow
        
        # 先还原包
        dotnet restore $projectFile
        if ($LASTEXITCODE -ne 0) {
            Write-Host "包还原失败!" -ForegroundColor Red
            exit 1
        }
        
        # 构建项目
        dotnet build $projectFile --configuration $Configuration --no-restore
        if ($LASTEXITCODE -eq 0) {
            Write-Host "构建完成!" -ForegroundColor Green
            Write-Host "可执行文件位置: bin\$Configuration\net6.0-windows\MarkdownToPdf.exe" -ForegroundColor Cyan
        } else {
            Write-Host "构建失败!" -ForegroundColor Red
            exit 1
        }
    }
    
    "run" {
        Write-Host "正在运行项目..." -ForegroundColor Yellow
        
        # 先构建
        dotnet build $projectFile --configuration $Configuration
        if ($LASTEXITCODE -ne 0) {
            Write-Host "构建失败!" -ForegroundColor Red
            exit 1
        }
        
        # 运行项目
        dotnet run --project $projectFile --configuration $Configuration
    }
    
    "publish" {
        Write-Host "正在发布项目..." -ForegroundColor Yellow
        $outputPath = "publish"
        
        dotnet publish $projectFile `
            --configuration Release `
            --output $outputPath `
            --self-contained false `
            --framework net6.0-windows
            
        if ($LASTEXITCODE -eq 0) {
            Write-Host "发布完成! 输出目录: $outputPath" -ForegroundColor Green
        } else {
            Write-Host "发布失败!" -ForegroundColor Red
            exit 1
        }
    }
    
    default {
        Write-Host "用法: build.ps1 [-Action <action>] [-Configuration <config>]" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "可用操作:" -ForegroundColor Yellow
        Write-Host "  restore  - 还原NuGet包"
        Write-Host "  clean    - 清理构建输出"
        Write-Host "  build    - 构建项目 (默认)"
        Write-Host "  run      - 构建并运行项目"
        Write-Host "  publish  - 发布项目"
        Write-Host ""
        Write-Host "可用配置:" -ForegroundColor Yellow
        Write-Host "  Debug    - 调试版本 (默认)"
        Write-Host "  Release  - 发布版本"
        Write-Host ""
        Write-Host "示例:" -ForegroundColor Yellow
        Write-Host "  .\build.ps1 -Action build"
        Write-Host "  .\build.ps1 -Action run"
        Write-Host "  .\build.ps1 -Action publish -Configuration Release"
    }
}

Write-Host ""
Write-Host "构建脚本执行完成!" -ForegroundColor Green 