@echo off

type into.txt
echo Adding Mica to Office...


echo Enabling Mica ( Microsoft.Office.UXPlatform.SVRefresh.UseMica)
REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\word" /v "Microsoft.Office.UXPlatform.SVRefresh.UseMica" /t REG_SZ /d true /f
REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\powerpoint" /v "Microsoft.Office.UXPlatform.SVRefresh.UseMica" /t REG_SZ /d true /f
REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\excel" /v "Microsoft.Office.UXPlatform.SVRefresh.UseMica" /t REG_SZ /d true /f
REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\onenote" /v "Microsoft.Office.UXPlatform.SVRefresh.UseMica" /t REG_SZ /d true /f

timeout /t 1 /nobreak >nul

echo Enabling BackStage Refresh (Microsoft.Office.UXPlatform.FluentSVBackstageRefresh)
REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\word" /v "Microsoft.Office.UXPlatform.FluentSVBackstageRefresh" /t REG_SZ /d true /f
REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\powerpoint" /v "Microsoft.Office.UXPlatform.FluentSVBackstageRefresh" /t REG_SZ /d true /f
REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\excel" /v "Microsoft.Office.UXPlatform.FluentSVBackstageRefresh" /t REG_SZ /d true /f
REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\onenote" /v "Microsoft.Office.UXPlatform.FluentSVBackstageRefresh" /t REG_SZ /d true /f

timeout /t 1 /nobreak >nul

Echo Enabling Menu Refresh V2 (Microsoft.Office.UXPlatform.FluentSVMenuRefresh_v2)
REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\word"  "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\powerpoint" /v "Microsoft.Office.UXPlatform.FluentSVMenuRefresh_v2" /t REG_SZ /d true /f
REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\powerpoint" /v "Microsoft.Office.UXPlatform.FluentSVMenuRefresh_v2" /t REG_SZ /d true /f
REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\excel" /v "Microsoft.Office.UXPlatform.FluentSVMenuRefresh_v2" /t REG_SZ /d true /f
REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\onenote" /v "Microsoft.Office.UXPlatform.FluentSVMenuRefresh_v2" /t REG_SZ /d true /f

timeout /t 1 /nobreak >nul

if %errorlevel% equ 0 (
    echo String value added to the registry successfully.
) else (
    echo Failed to add string value to the registry. Check for errors.
)

PAUSE