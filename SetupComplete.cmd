:: ---------------------------------------------------------------------------------------------------------------------------------------------------
:: Данный скрипт предназначен для поиска и передачи управления "по цепочке" другому (внешнему) BAT-файлу, который указан в переменной %MainScriptPath%
:: в этом файле. Имея название "SetupComplete.cmd" и находясь в папке"%HOMEDRIVE%\Windows\Setup\Scripts\", - куда данный файл следует поместить
:: заранее, - он автоматически будет запущен процессом "msoobe.exe" на завершающем этапе установки Windows 7 от имени встроенного Администратора.
:: Что-бы отключить логирование, необходимо закомментировать строку, которая устанавливает переменную %LogFilePath%.
:: ---------------------------------------------------------------------------------------------------------------------------------------------------

@Echo off
ChCp 866 >nul & Cls

SetLocal enableExtensions enableDelayedExpansion

	:: Установка значений переменных, необходимых для корректной работы этого скрипта.
	Set     "LogFilePath=%CD%\%~n0.log"
	Set   "SilentRunFlag=SKIP_CONFIRM"
	Set  "MainScriptPath=$_Res_for_Inst-Win7\Scripts\AutoSetup.bat"
	Set "ValidParentName=msoobe.exe"
	Set  "ValidFirstLine=:: This line is required to automatically run this script from "SetupComplete.cmd". Don't delete it^^^!"

	:: Защита от случайного запуска этого скрипта пользователем, путём проверки соответствия имени родительского процесса имени из %ValidParentName%.
	Call :LOG 150
	Call :CHECK_ProtectionFromStarting || Exit /b !ERRORLEVEL!
	Call :LOG 199

	:: Поиск валидного файла %MainScriptPath% на всех подключённых разделах и сохранение первого найденного в %runingScript%.
	Call :LOG 250
	Call :FIND_MainScriptFile || Exit /b !ERRORLEVEL!
	Call :LOG 299

	:: Передача управления найденному ранее и сохранённому в %runingScript% скрипту, который и будет выполнять основную работу.
	Call :LOG 350
	EndLocal & Call :RUN_runingScript "%runingScript%" "%SilentRunFlag%" "%LogFilePath%"
	Call :LOG 399

Exit

:: ---------------------------------------------------------------------------------------------------------------------------------------------------

:RUN_runingScript
	:: Установка рабочей папки для найденного скрипта из %runingScript% для дальнейшей корректной его работы.
	Cd /d "%~dp1"

	:: Запуск найденного скрипта из %runingScript%. А так же передача ему флага для "тихой" работы и пути к лог-файлу, куда он может писать ошибки.
	Call "%~1" "%~2" "%~3"

	:: Сохранение кода возврата, возвращённого отработавшим скриптом из %runingScript%.
	Set "returnСode=%ERRORLEVEL%"

	:: Повторная установка значений некоторым переменным. Т.к. для запуска скрипта из %runingScript% все переменные намеренно были стёрты.
	Set "runingScript=%~1"
	Set "LogFilePath=%~3"

	If "%returnСode%" gtr "0" ( Call :ERR 301 ) Else ( Call :LOG 351 )

	Exit /b 0

:: ---------------------------------------------------------------------------------------------------------------------------------------------------

:FIND_MainScriptFile
	SetLocal enableExtensions enableDelayedExpansion

		:: Перечисление всех подключённых разделов и поиск на них %MainScriptPath%.
		For /f "tokens=1,2 delims=: " %%A in ( 'WMIC LogicalDisk Get Caption^,Size ^| Find ":"' ) do For %%B in ( %%B ) do (
			Call :LOG 251 "%%A"

			:: Передача на валидацию найденного файла %MainScriptPath% и сохранение его в %runingScript%, в случае успеха.
			If exist "%%A:\%MainScriptPath%" (
				Call :LOG 252 "%%A"

				:: Получение первой строки из найденного файла %MainScriptPath% во временную переменную.
				Set /p tmpString=<"%%A:\%MainScriptPath%"
				If "!tmpString!" equ "%ValidFirstLine%" ( EndLocal & Set "runingScript=%%A:\%MainScriptPath%" & Exit /b 0 ) Else ( Call :ERR 201 )
			)
		)
		Call :ERR 202

	Exit /b 202

:CHECK_ProtectionFromStarting
	SetLocal enableExtensions enableDelayedExpansion

		:: Определение PID "текущего" CMD-файла.
		For /f "tokens=*" %%A in ( '
			Set "PPID=(Get-WmiObject Win32_Process -Filter ProcessId=$P).ParentProcessId" ^& ^
			Call powershell -NoLogo -NoProfile -Command "$P = $pid; $P = %%PPID%%; %%PPID%%"
		' ) do Set "this_processId=%%A"

		:: Определение PID родительского процесса по PID "текущего" CMD-файла.
		For /f "tokens=2 delims==" %%A in ( '
			WMIC PROCESS Where ^(processid^=%this_processId%^) Get ParentProcessId /value
		' ) do Set "parent_processId=%%A"

		:: Определение Имени родительского процесса по его же PID.
		For /f "tokens=2 delims==" %%A in ( '
			WMIC PROCESS Where ^(processid^=%parent_processId%^) Get Name /value
		' ) do Set "parent_processName=%%A"

		Call :LOG 151

		:: Генерация ошибки при "неверном" имени родительского процесса.
		If "%parent_processName%" neq "%ValidParentName%" (
			Call :ERR 101
			Exit /b 101
		)

	Exit /b 0

:: ---------------------------------------------------------------------------------------------------------------------------------------------------

:LOG
	SetLocal enableExtensions enableDelayedExpansion

		If "!LogFilePath!" equ "" Exit /b 0
		Set "d=%date:~6,4%/%date:~3,2%/%date:~0,2%" && Set "t=%time:~0,2%:%time:~3,2%:%time:~6,2%" && Set "t=!t: =0!"

		If "%~1" equ "150" ( (Echo.) & (Echo --------------------------------) & (Echo.) )>>"%LogFilePath%" 2>nul
		(
			<nul Set /p strTemp=[!d! !t!] %~1-
			If "%~1" equ "150" Echo BEG - Проверка соответствия имени родительского процесса "разрешённому".
			If "%~1" equ "151" Echo INF - PARENT_NAME="%parent_processName%"; VALID_NAME="%ValidParentName%"
			If "%~1" equ "199" Echo END - Успех: PARENT_NAME и VALID_NAME совпадают.
			If "%~1" equ "250" Echo BEG - Перечисление всех подключённых разделов и поиск на них "DRV:\%MainScriptPath%" до первого совпадения.
			If "%~1" equ "251" Echo INF - Поиск на DRV="%~2".
			If "%~1" equ "252" Echo INF - Обнаружен FILE_PATH="%~2:\%MainScriptPath%".
			If "%~1" equ "299" Echo END - Успех: Обнаруженный FILE_PATH валидный.
			If "%~1" equ "350" Echo BEG - Передача управления найденному скрипту: "%runingScript%".
			If "%~1" equ "351" Echo INF - Внешний скрипт "%runingScript%" завершил свою работу без ошибок: ERR="0".
			If "%~1" equ "399" Echo END - Завершение работы текущего скрипта.
		)>>"%LogFilePath%" 2>nul

	Cls & Exit /b 0

:ERR
	SetLocal enableExtensions enableDelayedExpansion

		If "!LogFilePath!" equ "" Exit /b 0
		Set "d=%date:~6,4%/%date:~3,2%/%date:~0,2%" && Set "t=%time:~0,2%:%time:~3,2%:%time:~6,2%" && Set "t=!t: =0!"

		(
			<nul Set /p strTemp=[!d! !t!] %~1-
			If "%~1" equ "101" Echo ERR - Ошибка: PARENT_NAME и VALID_NAME не совпадают. Выполнение скрипта прервано с кодом ошибки: "%~1".
			If "%~1" equ "201" Echo WAR - Предупреждение: Обнаруженный FILE_PATH не валидный. Поиск будет продолжен.
			If "%~1" equ "202" Echo ERR - Ошибка: Валидный FILE_PATH не обнаружен. Выполнение скрипта прервано с кодом ошибки: "%~1".
			If "%~1" equ "301" Echo WAR - Предупреждение: Скрипт "%runingScript%" завершил свою работу и вернул ошибку: ERR="%returnСode%".
		)>>"%LogFilePath%" 2>nul

	Cls & Exit /b 0

:: ---------------------------------------------------------------------------------------------------------------------------------------------------
