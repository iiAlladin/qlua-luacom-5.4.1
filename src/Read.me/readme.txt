Кратко: блок с пометкой “// LuaCOM Classfactory” реализует фабрику COM-классов, которая используется для «экспонирования» (Expose) Lua‑объекта как COM-сервиса вне процесса. В файле она представлена сигнатурой tLuaCOMClassFactory и используется в функциях luacom_ExposeObject / luacom_RevokeObject. Ниже — по пунктам как это работает и зачем.
Основная роль tLuaCOMClassFactory
•	Это реализующая IClassFactory обёртка вокруг Lua-объекта (внутри хранит IDispatch* object), чтобы Windows COM runtime мог создавать/получать экземпляры этого объекта по CLSID.
•	Методы:
•	QueryInterface, AddRef, Release — стандартный подсчёт ссылок и поддержка COM‑интерфейсов.
•	CreateInstance — возвращает указатель на требуемый интерфейс (обычно IUnknown/IDispatch) клиента, при необходимости делает QueryInterface на внутреннем объекте.
•	LockServer — используется COM для блокировки/отладки lifetime сервера.
Как фабрика используется в luacom.cpp
•	В luacom_ExposeObject (модифицированной версии) при необходимости «экспонировать» объект:
•	Если Lua работает in-process (luacom_runningInprocess(L)), простой inproc‑режим: объект сохраняется в реестре Lua как lightuserdata — нет регистрации класса в COM.
•	Если не in‑process — создаётся tLuaCOMClassFactory с внутренним IDispatch и регистрируется в COM через CoRegisterClassObject. Вызов возвращает cookie, который нужно передать в CoRevokeClassObject (см. luacom_RevokeObject) для отзыва фабрики.
•	luacom_RevokeObject вызывает CoRevokeClassObject(cookie) — снимает регистрацию фабрики в COM.
Регистрация в системном реестре (RegisterObject / UnRegisterObject)
•	luacom_RegisterObject:
•	Загружает и регистрирует типовую библиотеку (LoadTypeLibEx с REGKIND_REGISTER).
•	Извлекает LIBID, версию, CLSID CoClass и формирует ключи реестра (ProgID, VersionIndependentProgID, CLSID\{...}).
•	Записывает LocalServer32 = <ModulePath> [ScriptFile] [Arguments] — это указывает COM запускать текущий исполняемый файл как локальный сервер для CLSID.
•	Для скриповых/контролей добавляет дополнительные ключи (InprocServer32, ScriptFile, категории).
•	luacom_UnRegisterObject:
•	Загружает type lib (без регистрации), вызывает UnRegisterTypeLib, и удаляет соответствующие ключи реестра (ProgID, VersionIndependentProgID, CLSID\...).
Политика безопасности при регистрации/экспонировании
•	Новые функции check_registration_allowed и IsProcessElevated добавлены:
•	register/Expose требуют разрешения: либо luacom.config.allow_registration = true, либо процесс работает с повышенными правами (elevated), либо вызывающий передаёт корректный секретный token, совпадающий с module.config.registration_token.
•	Если проверка не проходит — операция блокируется и возвращается ошибка через luacom_APIerror.
•	Это защищает от несанкционированной записи в системный реестр и от регистрации class object.
Ошибки и исключения
•	Вся работа с COM обёрнута в try/catch(tLuaCOMException) — при ошибке вызывается luacom_APIerror / luacom_error, который помещает сообщение в module config и по настроенной политике либо бросает lua‑ошибку, либо сохраняет сообщение.
Жизненный цикл объектов и ссылки
•	tLuaCOMClassFactory держит ссылку на IDispatch (или на Lua↔COM объект). COM ожидает корректный подсчёт ссылок — AddRef/Release должны согласованно управлять временем жизни.
•	При CoRevokeClassObject COM перестаёт выдавать новые экземпляры; затем фабрика обычно освобождается (Release).
Где смотреть в коде
•	Создание и регистрация: luacom_ExposeObject.
•	Отзыв регистрации: luacom_RevokeObject.
•	Запись/удаление ключей реестра: luacom_RegisterObject / luacom_UnRegisterObject.
•	Политика разрешений: check_registration_allowed, IsProcessElevated.
•	Обработчик ошибок и сообщение пользователю: luacom_APIerror, luacom_error, luacom_err.