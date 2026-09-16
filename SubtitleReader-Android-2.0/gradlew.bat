@rem Gradle wrapper startup script for Windows
@if "%DEBUG%"=="" @echo off
set DIRNAME=%~dp0
if "%DIRNAME%"=="" set DIRNAME=.
set APP_BASE_NAME=%~n0
set APP_HOME=%DIRNAME%

set DEFAULT_JVM_OPTS="-Xmx64m" "-Xms64m"
set GRADLE_HOME=%DIRNAME%gradle\wrapper

if exist "%APP_HOME%\gradle\wrapper\gradle-wrapper.jar" (
    set CLASSPATH=%APP_HOME%\gradle\wrapper\gradle-wrapper.jar
) else (
    echo Gradle wrapper jar not found, downloading...
    powershell -Command "& { Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/gradle/gradle/v8.5.0/gradle/wrapper/gradle-wrapper.jar' -OutFile '%APP_HOME%\gradle\wrapper\gradle-wrapper.jar' }"
    if exist "%APP_HOME%\gradle\wrapper\gradle-wrapper.jar" (
        set CLASSPATH=%APP_HOME%\gradle\wrapper\gradle-wrapper.jar
    ) else (
        echo Failed to download gradle-wrapper.jar
        echo Please open this project in Android Studio instead.
        exit /b 1
    )
)

"%JAVA_HOME%/bin/java.exe" %DEFAULT_JVM_OPTS% -classpath "%CLASSPATH%" org.gradle.wrapper.GradleWrapperMain %*
