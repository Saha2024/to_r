
#Область СлужебныйПрограммныйИнтерфейс

#Область Авторизация

// Выполняет расшифровку маркеров.
// 
// Параметры:
// 	Оповещение - ОписаниеОповещения - содержит описание процедуры, которая будет выполнена после расшифровки маркеров
// 	             со следующими параметрами:
// 	               * Результат - Структура:
// 	                   ** КлючиСинхронизации - см. СинхронизацияЭДОКлиентСервер.НовыеКлючиСинхронизации
// 	                   ** КонтекстДиагностики - см. ОбработкаНеисправностейБЭДКлиент.НовыйКонтекстДиагностики
// 	                   ** ПаролиСертификатов - см. КриптографияБЭДКлиент.НовыеПаролиСертификатов
// 	                   ** ОбработанныеУчетныеЗаписи - Соответствие
// 	               * ДополнительныеПараметры - Произвольный - значение, которое было указано при создании
// 	                 объекта ОписаниеОповещения.
// 	КлючиСинхронизации - см. СинхронизацияЭДОКлиентСервер.НовыеКлючиСинхронизации
// 	КонтекстДиагностики - см. ОбработкаНеисправностейБЭДКлиент.НовыйКонтекстДиагностики
// 	ПаролиСертификатов - см. КриптографияБЭДКлиент.НовыеПаролиСертификатов
// 	ОбработанныеУчетныеЗаписи - Соответствие
Процедура ПолучитьПараметрыАутентификации(Оповещение, КлючиСинхронизации, КонтекстДиагностики,
	ПаролиСертификатов = Неопределено, ОбработанныеУчетныеЗаписи = Неопределено) Экспорт
	
	Контекст = НовыйКонтекстПолученияПараметровАутентификации();
	
	ДанныеСертификатов = Новый Соответствие;
	Для Каждого КлючИЗначение Из КлючиСинхронизации Цикл
		МаркерыПоСертификатам = Новый Соответствие;
		Для Каждого ОписаниеСертификата Из КлючИЗначение.Значение.Сертификаты Цикл
			ДанныеСертификатов.Вставить(ОписаниеСертификата.Ссылка, ОписаниеСертификата);
			Если ОписаниеСертификата.Свойство("МаркерЗашифрованный") Тогда
				МаркерыПоСертификатам.Вставить(ОписаниеСертификата.Ссылка, ОписаниеСертификата.МаркерЗашифрованный);
			КонецЕсли;
		КонецЦикла;
		Контекст.ЗашифрованныеМаркеры.Вставить(КлючИЗначение.Ключ, МаркерыПоСертификатам);
	КонецЦикла;
	
	Контекст.КлючиСинхронизации = КлючиСинхронизации;
	Контекст.ДанныеСертификатов = ДанныеСертификатов;
	Контекст.КонтекстДиагностики = КонтекстДиагностики;
	Если ПаролиСертификатов <> Неопределено Тогда
		Контекст.ПаролиСертификатов = ПаролиСертификатов;
	КонецЕсли;
	Если ОбработанныеУчетныеЗаписи = Неопределено Тогда
		Контекст.ОбработанныеУчетныеЗаписи = Новый Соответствие;
	Иначе
		Контекст.ОбработанныеУчетныеЗаписи = ОбработанныеУчетныеЗаписи;
	КонецЕсли;
	Контекст.ОповещениеОЗавершении = Оповещение;
	
	Обработчики = Новый Массив;
	Обработчики.Добавить(Новый ОписаниеОповещения("РасшифроватьМаркер", ЭтотОбъект));
	
	МассивКлючейСинхронизации = Новый Массив;
	Для Каждого КлючИЗначение Из КлючиСинхронизации Цикл
		МассивКлючейСинхронизации.Добавить(КлючИЗначение.Значение);
	КонецЦикла;
	ОповещениеОЗавершенииСлужебный = Новый ОписаниеОповещения("ПослеРасшифровкиМаркеров", ЭтотОбъект, Контекст);
	ОбщегоНазначенияБЭДКлиент.АсинхронныйЦиклНачать(ОповещениеОЗавершенииСлужебный, МассивКлючейСинхронизации,
		Обработчики, Контекст);

КонецПроцедуры

#КонецОбласти

#Область РегистрацияУчетныхЗаписей

// Возвращает параметры регистрации в сервисе 1С-ЭДО.
// 
// Возвращаемое значение:
// 	Структура - Описание:
// * ИдентификаторОрганизации - Строка
// * СертификатыПодписейОрганизации - Массив из СправочникСсылка.СертификатыКлючейЭлектроннойПодписиИШифрования
// * КодНалоговогоОргана - Строка
// * ЭлектроннаяПочта - Строка
// * СертификатКриптографии - СправочникСсылка.СертификатыКлючейЭлектроннойПодписиИШифрования
// * АдресОрганизацииЗначение - Строка
// * АдресОрганизации - Строка
// * ОператорЭДО - Строка
// * Организация - ОпределяемыйТип.Организация
// * Доверенности - КлючИЗначение:
//    ** Ключ - СправочникСсылка.СертификатыКлючейЭлектроннойПодписиИШифрования
//    ** Значение - СправочникСсылка.МашиночитаемыеДоверенностиОрганизаций
Функция НовыеПараметрыРегистрацииВСервисе1СЭДО() Экспорт
	
	ПараметрыРегистрации = Новый Структура;
	ПараметрыРегистрации.Вставить("Организация"                    , Неопределено);
	ПараметрыРегистрации.Вставить("ОператорЭДО"                    , "");
	ПараметрыРегистрации.Вставить("АдресОрганизации"               , "");
	ПараметрыРегистрации.Вставить("АдресОрганизацииЗначение"       , "");
	ПараметрыРегистрации.Вставить("СертификатКриптографии"         ,
		ПредопределенноеЗначение("Справочник.СертификатыКлючейЭлектроннойПодписиИШифрования.ПустаяСсылка"));
	ПараметрыРегистрации.Вставить("ЭлектроннаяПочта"               , "");
	ПараметрыРегистрации.Вставить("КодНалоговогоОргана"            , "");
	ПараметрыРегистрации.Вставить("СертификатыПодписейОрганизации" , Новый Массив);
	ПараметрыРегистрации.Вставить("ИдентификаторОрганизации"       , "");
	ПараметрыРегистрации.Вставить("ОтборСертификатов"			   , Новый Массив);
	ПараметрыРегистрации.Вставить("Доверенности"				   , Новый Соответствие);
	
	Возврат ПараметрыРегистрации;
	
КонецФункции

// Получает состояние заявки на регистрацию в сервисе 1С-ЭДО.
// 
// Параметры:
// 	ИдентификаторыЗаявок - Массив из Строка
// 	Оповещение - ОписаниеОповещения - содержит описание процедуры, которая будет вызвана после отправки данных
// 	                     со следующими параметрами:
// 	                     Результат - см. СервисЭДО.ПолучитьСостояниеРегистрацииВСервисе1СЭДО
// 	                     ДополнительныеПараметры - Произвольный - значение, которое было указано при создании 
// 	                     объекта ОписаниеОповещения.
// 	ИдентификаторФормы - УникальныйИдентификатор - идентификатор формы, во временное хранилище которой надо поместить
//  результат выполнения процедуры.
Процедура ПолучитьСостояниеРегистрацииВСервисе1СЭДО(ИдентификаторыЗаявок, Оповещение, ИдентификаторФормы) Экспорт

	ОперацияПолученияСостояния = СервисЭДОВызовСервера.НачатьПолучениеСостоянияРегистрацииВСервисе1СЭДО(
		ИдентификаторыЗаявок, ИдентификаторФормы);

	ПараметрыОжидания = ДлительныеОперацииКлиент.ПараметрыОжидания(Неопределено);
	ПараметрыОжидания.ВыводитьОкноОжидания = Ложь;

	ДлительныеОперацииКлиент.ОжидатьЗавершение(ОперацияПолученияСостояния, Оповещение, ПараметрыОжидания);

КонецПроцедуры

// Подготавливает для подписания регистрационные данные в сервисе 1С-ЭДО.
// 
// Параметры:
// 	ДанныеПакета - см. СервисЭДО.ДанныеДляРегистрационногоПакета1СЭДО
// Возвращаемое значение:
// 	Структура - Описание:
// * ПредставлениеНабора - Строка -
// * СписокПредставлений - СписокЗначений -
// * КонтекстОперации - Неопределено -
// * БезПодтверждения - Булево -
// * ЗаголовокДанных - Строка -
// * ПоказатьКомментарий - Булево -
// * НаборДанных - Массив -
// * ОтборСертификатов - Массив -
// * Операция - Строка -
Функция ПодготовитьДанныеНаПодпись(ДанныеПакета, СертификатыПодписейОрганизации = Неопределено) Экспорт
	
	ДанныеНаПодпись = Новый Структура;
	
	Операция = НСтр("ru = 'Подписание регистрационных данных на подключение к сервису 1С-ЭДО'");
	ДанныеНаПодпись.Вставить("Операция", Операция);
	
	Если СертификатыПодписейОрганизации <> Неопределено Тогда
		ДанныеНаПодпись.Вставить("ОтборСертификатов", СертификатыПодписейОрганизации);
	Иначе
		ДанныеНаПодпись.Вставить("ОтборСертификатов",
			ОбщегоНазначенияКлиентСервер.ЗначениеВМассиве(ДанныеПакета.ДобавляемыйСертификат))
	КонецЕсли;
	
	ДанныеНаПодпись.Вставить("НаборДанных", Новый Массив);
	ДанныеНаПодпись.НаборДанных.Добавить(Новый Структура("Данные",
		ДанныеПакета.ДвоичныеДанныеСоглашенияНаПодключениеЭДО));
	
	ДанныеНаПодпись.НаборДанных.Добавить(Новый Структура("Данные", ДанныеПакета.ДвоичныеДанныеДляОператораЭДО));
	
	ДанныеНаПодпись.Вставить("ПоказатьКомментарий", Ложь);
	ДанныеНаПодпись.Вставить("ЗаголовокДанных", "");
	ДанныеНаПодпись.Вставить("БезПодтверждения", Истина);
	ДанныеНаПодпись.Вставить("КонтекстОперации", Неопределено);
	
	СписокПредставлений = Новый СписокЗначений;
	
	Контекст = Новый Структура;
	Контекст.Вставить("РеквизитыПакета", ДанныеПакета.РеквизитыПакета);
	Контекст.Вставить("ДвоичныеДанныеДляОператораЭДО", ДанныеПакета.ДвоичныеДанныеДляОператораЭДО);
	
	Представление = НСтр("ru = 'Пользовательское соглашение на подключение к сервису 1С-ЭДО'");
	ОписаниеПредставленияЗаявления = Новый ОписаниеОповещения("ОбработатьПредставлениеЗаявленияНаПодключениеКЭДО",
		СервисЭДОКлиент, Контекст);
	СписокПредставлений.Добавить(ОписаниеПредставленияЗаявления, Представление);
	
	Представление = НСтр("ru = 'Данные для регистрации учетной записи у оператора ЭДО'");
	ОписаниеПредставленияДанных = Новый ОписаниеОповещения("ОбработатьПредставлениеДанныхДляОператораЭДО",
		СервисЭДОКлиент, Контекст);
	СписокПредставлений.Добавить(ОписаниеПредставленияДанных, Представление);
	
	ДанныеНаПодпись.Вставить("СписокПредставлений", СписокПредставлений);
	ДанныеНаПодпись.Вставить("ПредставлениеНабора",
		НСтр("ru = 'Пользовательское соглашение, Данные для регистрации у оператора ЭДО'"));
	
	Возврат ДанныеНаПодпись;
	
КонецФункции

// Открывает регистрационные данные в сервисе 1С-ЭДО.
// 
// Параметры:
// 	РеквизитыПакета - см. СервисЭДО.ДанныеДляРегистрационногоПакета1СЭДО.РеквизитыПакета
// 	ДвоичныеДанныеДляОператораЭДО - см. СервисЭДО.ДанныеДляРегистрационногоПакета1СЭДО.ДвоичныеДанныеДляОператораЭДО
Процедура ОткрытьДанныеДляРегистрацииУОператораЭлектронногоДокументооборота(РеквизитыПакета, ДвоичныеДанныеДляОператораЭДО) Экспорт
	
	ШаблонИмениВременногоФайла = НСтр("ru = 'Данные для регистрации у оператора ЭДО_%1'");
	ИмяВременногоФайла = СтроковыеФункцииКлиентСервер.ПодставитьПараметрыВСтроку(ШаблонИмениВременногоФайла,
		РеквизитыПакета.УникальныйИдентификаторЗаявки1СЭДО);
	ИмяВременногоФайлаСРасширением = ОбщегоНазначенияКлиентСервер.ПолучитьИмяСРасширением(ИмяВременногоФайла, "xml");
	
	СсылкаНаДвоичныеДанныеФайла = ПоместитьВоВременноеХранилище(ДвоичныеДанныеДляОператораЭДО, Новый УникальныйИдентификатор);
	
	ФайловаяСистемаКлиент.ОткрытьФайл(СсылкаНаДвоичныеДанныеФайла,, ИмяВременногоФайлаСРасширением);
	
КонецПроцедуры

// Открывает пользовательское соглашение сервиса 1С-ЭДО.
// 
Процедура ОткрытьПользовательскоеСоглашение1СЭДО() Экспорт
	
	ФайловаяСистемаКлиент.ОткрытьНавигационнуюСсылку("https://portal.1c.ru/applications/30/licenseAgreement");
	
КонецПроцедуры

#КонецОбласти

#Область Уведомления

// Возвращает настройки уведомлений учетной записи в сервисе 1С-ЭДО.
// 
// Параметры:
// 	КлючСинхронизации - см. СинхронизацияЭДОКлиентСервер.НовыйКлючСинхронизации
// 	КонтекстДиагностики - см. ОбработкаНеисправностейБЭДКлиент.НовыйКонтекстДиагностики
// 	ОповещениеЗавершении - ОписаниеОповещения - содержит описание процедуры, которая будет вызвана после получения
// 	                     параметров уведомлений со следующими параметрами:
// 	                     Результат - см. СервисЭДО.ПолучитьПараметрыУведомлений
// 	                     ДополнительныеПараметры - Произвольный - значение, которое было указано при создании 
// 	                     объекта ОписаниеОповещения.
// 	Форма - ФормаКлиентскогоПриложения - форма, в хранилище которой нужно поместить результат.
// 	ПаролиСертификатов - см. КриптографияБЭДКлиент.НовыеПаролиСертификатов
Процедура ПолучитьПараметрыУведомлений(КлючСинхронизации, КонтекстДиагностики, ОповещениеЗавершении,
	Форма = Неопределено, ПаролиСертификатов = Неопределено) Экспорт
	
	Контекст = Новый Структура;
	Контекст.Вставить("КлючСинхронизации", КлючСинхронизации);
	Контекст.Вставить("КонтекстДиагностики", КонтекстДиагностики);
	Контекст.Вставить("ПараметрыУведомлений", Неопределено);
	Контекст.Вставить("ОповещениеОЗавершении", ОповещениеЗавершении);
	Контекст.Вставить("Форма", Форма);
	
	Оповещение = Новый ОписаниеОповещения(
		"ПолучитьПараметрыУведомленийПослеПолученияПараметровАутентификации", ЭтотОбъект, Контекст);
	КлючиСинхронизации = СинхронизацияЭДОКлиентСервер.НовыеКлючиСинхронизации();
	КлючиСинхронизации.Вставить(КлючСинхронизации.ИдентификаторУчетнойЗаписи, КлючСинхронизации);
	ПолучитьПараметрыАутентификации(Оповещение, КлючиСинхронизации, КонтекстДиагностики, ПаролиСертификатов);
	
КонецПроцедуры

// Возвращает настройки уведомлений учетной записи в сервисе 1С-ЭДО.
// 
// Параметры:
// 	КлючСинхронизации - см. СинхронизацияЭДОКлиентСервер.НовыйКлючСинхронизации
// 	ОписаниеОповещения - ОписаниеОповещения - содержит описание процедуры, которая будет вызвана после отправки данных
// 	                     со следующими параметрами:
// 	                     Результат - см. СервисЭДО.ПолучитьПараметрыУведомлений
// 	                     ДополнительныеПараметры - Произвольный - значение, которое было указано при создании 
// 	                     объекта ОписаниеОповещения.
// 	Форма - ФормаКлиентскогоПриложения - форма, во временное хранилище которой надо поместить результат
//  выполнения процедуры.
Процедура НачатьПолучениеПараметровУведомлений(КлючСинхронизации,
	ОписаниеОповещения, Форма = Неопределено) Экспорт

	ПараметрыОжидания = ДлительныеОперацииКлиент.ПараметрыОжидания(Форма);
	ПараметрыОжидания.ВыводитьОкноОжидания = Ложь;
	ПараметрыОжидания.ОповещениеПользователя.Показать = Ложь;
	ПараметрыОжидания.ВыводитьСообщения = Истина;
	
	ИдентификаторФормы = ?(Форма = Неопределено, Новый УникальныйИдентификатор(), Форма.УникальныйИдентификатор);
	ДлительнаяОперация = СервисЭДОВызовСервера.НачатьПолучениеПараметровУведомлений(
		КлючСинхронизации, ИдентификаторФормы);
	
	ДлительныеОперацииКлиент.ОжидатьЗавершение(ДлительнаяОперация,
		ОписаниеОповещения,
		ПараметрыОжидания);
		
КонецПроцедуры

// Обновляет настройки уведомлений учетных записей в сервисе 1C-ЭДО.
// 
// Параметры:
// 	ПараметрыУведомлений - Строка
// 	КлючСинхронизации - см. СинхронизацияЭДОКлиентСервер.НовыйКлючСинхронизации
// 	КонтекстДиагностики - см. ОбработкаНеисправностейБЭДКлиент.НовыйКонтекстДиагностики
// 	ОписаниеОповещения - ОписаниеОповещения - содержит описание процедуры, которая будет вызвана после отправки данных
// 	                     со следующими параметрами:
// 	                     Результат - см. СервисЭДО.ОбновитьПараметрыУведомленийВСервисеНастроек
// 	                     ДополнительныеПараметры - Произвольный - значение, которое было указано при создании 
// 	                     объекта ОписаниеОповещения.
// 	Форма - ФормаКлиентскогоПриложения - форма, во временное хранилище которой надо поместить результат
// выполнения процедуры.
// 	ПаролиСертификатов - см. КриптографияБЭДКлиент.НовыеПаролиСертификатов
Процедура ОбновитьПараметрыУведомлений(ПараметрыУведомлений, КлючСинхронизации, КонтекстДиагностики, ОписаниеОповещения,
	Форма = Неопределено, ПаролиСертификатов = Неопределено) Экспорт
	
	Контекст = Новый Структура;
	Контекст.Вставить("КлючСинхронизации", КлючСинхронизации);
	Контекст.Вставить("КонтекстДиагностики", КонтекстДиагностики);
	Контекст.Вставить("ПараметрыУведомлений", ПараметрыУведомлений);
	Контекст.Вставить("ОповещениеОЗавершении", ОписаниеОповещения);
	Контекст.Вставить("Форма", Форма);
	Контекст.Вставить("Успех", Ложь);
	
	Оповещение = Новый ОписаниеОповещения(
		"ОбновитьПараметрыУведомленийПослеПолученияПараметровАутентификации", ЭтотОбъект, Контекст);
	КлючиСинхронизации = СинхронизацияЭДОКлиентСервер.НовыеКлючиСинхронизации();
	КлючиСинхронизации.Вставить(КлючСинхронизации.ИдентификаторУчетнойЗаписи, КлючСинхронизации);
	ПолучитьПараметрыАутентификации(Оповещение, КлючиСинхронизации, КонтекстДиагностики, ПаролиСертификатов);
		
КонецПроцедуры

#КонецОбласти

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

// Расшифровывает маркер.
// 
// Параметры:
// 	КлючСинхронизации - см. СинхронизацияЭДОКлиентСервер.НовыйКлючСинхронизации
// 	Контекст - см. НовыйКонтекстПолученияПараметровАутентификации
Процедура РасшифроватьМаркер(КлючСинхронизации, Контекст) Экспорт
	
	ИдентификаторУчетнойЗаписи = КлючСинхронизации.ИдентификаторУчетнойЗаписи;
	Сертификаты = КлючСинхронизации.Сертификаты;

	Если НЕ (ТипЗнч(Сертификаты) = Тип("Массив")
			И ЗначениеЗаполнено(ИдентификаторУчетнойЗаписи)) Тогда
		ОбщегоНазначенияБЭДКлиент.АсинхронныйЦиклВыполнитьСледующийОбработчик(Неопределено, Контекст);
		Возврат
	КонецЕсли;

	МассивСертификатов = Новый Массив;
	
	Если ЗначениеЗаполнено(КлючСинхронизации.ВыбранныйСертификат) Тогда
		МассивСертификатов.Добавить(КлючСинхронизации.ВыбранныйСертификат);
	Иначе
		Для Каждого Сертификат Из Сертификаты Цикл
			Если ЗначениеЗаполнено(Сертификат.МаркерРасшифрованный) Тогда
				МассивСертификатов = Новый Массив;
				Прервать;
			Иначе
				
				// Если хранилище данных сертификата ничего не содержит, то такой сертификат не добавляется в массив
				СертификатыДляАнализа = ОбщегоНазначенияКлиентСервер.ЗначениеВМассиве(Сертификат.Ссылка);
				СвойстваСертификатов = КриптографияБЭДВызовСервера.СвойстваСертификатов(СертификатыДляАнализа);
				ДанныеТекущегоСертификата = СвойстваСертификатов[Сертификат.Ссылка].ДанныеСертификата;
				Если Не ЗначениеЗаполнено(ДанныеТекущегоСертификата) Тогда
					Продолжить;
				КонецЕсли;
				
				Если Сертификат.ПарольПолучен Тогда
					// Авторизоваться на сервере оператора можно любым зарегистрированным в соглашении сертификатом,
					// поэтому, если есть несколько доступных для авторизации сертификатов и среди них есть хотя бы
					// один с сохраненным (в сертификате или сеансе) паролем, то вернем именно его, чтобы не открывать
					// диалог выбора сертификата.
					МассивСертификатов = Новый Массив;
					МассивСертификатов.Добавить(Сертификат.Ссылка);
					Прервать;
				Иначе
					МассивСертификатов.Добавить(Сертификат.Ссылка);
				КонецЕсли;
			КонецЕсли;
		КонецЦикла;
	КонецЕсли;
	
	// Если массив сертификатов пустой, значит либо уже есть расшифрованный маркер, либо нет сертификатов,
	// в обоих случаях переходим к обработке следующей Настройки ЭДО.
	Если МассивСертификатов.Количество() > 0 Тогда

		// Если пароль к сертификату временно сохранен, то получим его и укажем в качестве контекста операции.
		Если ЗначениеЗаполнено(КлючСинхронизации.ВыбранныйСертификат) Тогда
			КонтекстОперации = Контекст.ПаролиСертификатов[КлючСинхронизации.ВыбранныйСертификат];
		КонецЕсли;
		
		КриптографияБЭДКлиент.УстановитьПаролиСертификатов(Контекст.ДанныеСертификатов);
		
		ОписаниеОповещения = Новый ОписаниеОповещения("ПослеРасшифровкиМаркера", ЭтотОбъект, Контекст);
		Маркер = Новый ОписаниеОповещения("ПолучитьДанныеЗашифрованногоМаркера", ЭтотОбъект, Контекст);
		ОписаниеДанных = Новый Структура;
		ОписаниеДанных.Вставить("Данные", Маркер);
		ОписаниеДанных.Вставить("ЗаголовокДанных", "");
		ОписаниеДанных.Вставить("КонтекстОперации", КонтекстОперации);
		ОписаниеДанных.Вставить("ОтборСертификатов", МассивСертификатов);
		ОписаниеДанных.Вставить("Операция", НСтр("ru = 'Аутентификация в сервисе ЭДО'"));
		ОписаниеДанных.Вставить("БезПодтверждения",  Истина);
		ОписаниеДанных.Вставить("ЭтоАутентификация", Истина);
		ОписаниеДанных.Вставить("РазрешитьЗапоминатьПароль", Истина);
		ОписаниеДанных.Вставить("ПрекратитьВыполнение", Истина);
		
		СинхронизацияЭДОКлиентСервер.УстановитьТекущуюУчетнуюЗапись(Контекст.КонтекстДиагностики,
			ИдентификаторУчетнойЗаписи);
		
		КриптографияБЭДКлиент.Расшифровать(ОписаниеДанных, Контекст.КонтекстДиагностики,, ОписаниеОповещения,
			Контекст.ПаролиСертификатов);
	Иначе
		ОбщегоНазначенияБЭДКлиент.АсинхронныйЦиклВыполнитьСледующийОбработчик(Неопределено, Контекст);
	КонецЕсли;
	
КонецПроцедуры

Процедура ПолучитьДанныеЗашифрованногоМаркера(Результат, Контекст) Экспорт
	
	// Результат - структура - ОписаниеДанных которая была передана в метод Расшифровать(),
	// дополненная параметром ВыбранныйСертификат на стороне БСП:
	Если Результат.ОписаниеДанных.Свойство("ВыбранныйСертификат") Тогда
		
		Контекст.Вставить("ВыбранныйСертификат", Результат.ОписаниеДанных.ВыбранныйСертификат);
		Контекст.Вставить("РезультатВыполнения", Результат);

		ПолучитьДанныеЗашифрованногоМаркераПродолжить(Контекст);
	Иначе
		ОбщегоНазначенияБЭДКлиент.АсинхронныйЦиклВыполнитьСледующийОбработчик(Неопределено, Контекст);
	КонецЕсли;
	
КонецПроцедуры

Процедура ПолучитьДанныеЗашифрованногоМаркераПродолжить(Контекст, ПредложитьПодключитьИнтернетПоддержку = Истина)
	
	КлючСинхронизации = ОбщегоНазначенияБЭДКлиент.АсинхронныйЦиклТекущийЭлементДанных(Контекст);
	
	МаркерыПоСертификатам = Контекст.ЗашифрованныеМаркеры[КлючСинхронизации.ИдентификаторУчетнойЗаписи];
	Если МаркерыПоСертификатам = Неопределено Тогда
		МаркерЗашифрованный = Неопределено;
	Иначе
		МаркерЗашифрованный = МаркерыПоСертификатам[Контекст.ВыбранныйСертификат.Ссылка];
	КонецЕсли;
	
	ДанныеСертификата = Контекст.ДанныеСертификатов[Контекст.ВыбранныйСертификат.Ссылка];
	
	КонтекстДиагностики = Контекст.КонтекстДиагностики;
	Если МаркерЗашифрованный = Неопределено Тогда
		
		ДвоичныеДанныеСертификата = ?(ДанныеСертификата = Неопределено, Неопределено, ДанныеСертификата.ДвоичныеДанные);
		ПараметрыАвторизации = Неопределено;
		МаркерЗашифрованный = СервисЭДОВызовСервера.Авторизоваться(КлючСинхронизации, КонтекстДиагностики,
			ДвоичныеДанныеСертификата, ПараметрыАвторизации);
		
		Если ПредложитьПодключитьИнтернетПоддержку И ПараметрыАвторизации <> Неопределено
			И ПараметрыАвторизации.ТребуетсяПодключениеИнтернетПоддержки Тогда
			
			ИнтернетПоддержкаПользователейКлиент.ПодключитьИнтернетПоддержкуПользователей(
				Новый ОписаниеОповещения("ПолучитьДанныеЗашифрованногоМаркераАутентификация", ЭтотОбъект, Контекст),
				ЭтотОбъект);
			Возврат;
		КонецЕсли;
		
	КонецЕсли;
	
	Контекст.КонтекстДиагностики = КонтекстДиагностики;
		
	Параметры = Новый Структура;
	Параметры.Вставить("Данные", МаркерЗашифрованный);
	
	Если МаркерЗашифрованный = Неопределено Тогда
		ВидОперации = НСтр("ru = 'Аутентификация в сервисе оператора электронного документооборота'");
		ТекстОшибкиШаблон = НСтр("ru = 'Не удалось получить зашифрованный маркер.'") + Символы.ПС 
			+ НСтр("ru = 'Сертификат электронной подписи: %1.'") + Символы.ПС
			+ НСтр("ru = 'Способ обмена: %2.'");
		ТекстОшибки = СтрШаблон(ТекстОшибкиШаблон, Контекст.ВыбранныйСертификат.Ссылка,
			КлючСинхронизации.СпособОбмена);
		
		ОбработкаНеисправностейБЭДВызовСервера.ОбработатьОшибку(ВидОперации,
				ОбщегоНазначенияБЭДКлиентСервер.ПодсистемыБЭД().ОбменСКонтрагентами, ТекстОшибки);
			
		ОписаниеОшибки = НСтр("ru = 'Не удалось пройти аутентификацию в сервисе оператора электронного документооборота.'");
		Параметры.Вставить("ОписаниеОшибки", ОписаниеОшибки + Символы.ПС + НСтр("ru = 'Подробнее см. в журнале регистрации.'"));
	КонецЕсли;
	
	ВыполнитьОбработкуОповещения(Контекст.РезультатВыполнения.Оповещение, Параметры);
	
КонецПроцедуры

Процедура ПослеРасшифровкиМаркера(Результат, Контекст) Экспорт
	
	ТекущийЭлементДанных = ОбщегоНазначенияБЭДКлиент.АсинхронныйЦиклТекущийЭлементДанных(Контекст);
	ИдентификаторУчетнойЗаписи = ТекущийЭлементДанных.ИдентификаторУчетнойЗаписи;
	Если Результат.Свойство("ВыбранныйСертификат") Тогда
		ТекущийЭлементДанных.ВыбранныйСертификат = Результат.ВыбранныйСертификат.Ссылка;
	КонецЕсли;
	Если Результат.Свойство("РасшифрованныеДанные") Тогда
		Если ТипЗнч(Результат.РасшифрованныеДанные) = Тип("Строка") Тогда
			РасшифрованныеДанные = ПолучитьИзВременногоХранилища(Результат.РасшифрованныеДанные);
		Иначе
			РасшифрованныеДанные = Результат.РасшифрованныеДанные;
		КонецЕсли;
		ТекущийЭлементДанных.МаркерРасшифрованный = РасшифрованныеДанные;
	
		Если Результат.Свойство("ПаролиСертификатов") Тогда
			Контекст.ПаролиСертификатов = Результат.ПаролиСертификатов;
		КонецЕсли;
	
	КонецЕсли;
	
	ОбработатьОшибкуУчетнойЗаписи(Результат, ИдентификаторУчетнойЗаписи, Контекст);
	
	ОбщегоНазначенияБЭДКлиент.АсинхронныйЦиклВыполнитьСледующийОбработчик(Истина, Контекст);
	
КонецПроцедуры

Процедура ПослеРасшифровкиМаркеров(Результат, Контекст) Экспорт
	
	КриптографияБЭДКлиент.УстановитьПаролиСертификатов(Контекст.ДанныеСертификатов, Ложь);
	
	КлючиСинхронизации = СинхронизацияЭДОКлиентСервер.НовыеКлючиСинхронизации();
	Для Каждого КлючИЗначение Из Результат.КлючиСинхронизации Цикл
		Если ЗначениеЗаполнено(КлючИЗначение.Значение.МаркерРасшифрованный) Тогда
			КлючиСинхронизации.Вставить(КлючИЗначение.Ключ, КлючИЗначение.Значение);
		КонецЕсли;
	КонецЦикла;
	ПараметрыАутентификации = Новый Структура;
	ПараметрыАутентификации.Вставить("КлючиСинхронизации", КлючиСинхронизации);
	ПараметрыАутентификации.Вставить("КонтекстДиагностики", Результат.КонтекстДиагностики);
	ПараметрыАутентификации.Вставить("ПаролиСертификатов", Результат.ПаролиСертификатов);
	ПараметрыАутентификации.Вставить("ОбработанныеУчетныеЗаписи", Результат.ОбработанныеУчетныеЗаписи);
	
	ВыполнитьОбработкуОповещения(Контекст.ОповещениеОЗавершении, ПараметрыАутентификации);
	
КонецПроцедуры

Процедура ОбработатьОшибкуУчетнойЗаписи(Результат, ИдентификаторУчетнойЗаписи, Контекст)
	
	ОбработаннаяУчетнаяЗапись = ДиагностикаЭДОКлиентСервер.НовоеОписаниеОбработаннойУчетнойЗаписи();
	ОбработаннаяУчетнаяЗапись.Идентификатор = ИдентификаторУчетнойЗаписи;
	ОбработаннаяУчетнаяЗапись.ОтправкаПолучениеВыполнялись = Результат.Успех;
	
	Если Не Результат.Успех Тогда
		ПользовательОтказалсяОтОперации = Результат.Отказ;
		Если ПользовательОтказалсяОтОперации Тогда
			ОбработаннаяУчетнаяЗапись.ПользовательОтказалсяОтОперации = Истина;
		Иначе
			ИнформацияОбОшибке = ОбработаннаяУчетнаяЗапись.ИнформацияОбОшибке;
			ИнформацияОбОшибке.КраткоеПредставление = НСтр("ru = 'Ошибка при расшифровке маркера'");
			Если Результат.Свойство("ОписаниеОшибки") Тогда
				ИнформацияОбОшибке.ПодробноеПредставление.Текст = Результат.ОписаниеОшибки;
			КонецЕсли;
		КонецЕсли;
	КонецЕсли;
	Контекст.ОбработанныеУчетныеЗаписи.Вставить(ОбработаннаяУчетнаяЗапись.Идентификатор, ОбработаннаяУчетнаяЗапись);
	
КонецПроцедуры

Процедура ПолучитьДанныеЗашифрованногоМаркераАутентификация(ПараметрыАутентификации, Контекст = Неопределено) Экспорт
	
	Если ТипЗнч(ПараметрыАутентификации) <> Тип("Структура") Тогда
		// Пользователь отказался от ввода логина и пароля.
		ПолучитьДанныеЗашифрованногоМаркераПродолжить(Контекст, Ложь);
		Возврат;
	КонецЕсли;
	
	ПолучитьДанныеЗашифрованногоМаркераПродолжить(Контекст);
	
КонецПроцедуры

// Возвращает структуру, используемую при получении параметров аутентификации.
// 
// Возвращаемое значение:
// 	Структура - Описание:
// * ОбработанныеУчетныеЗаписи - Соответствие
// * ЗашифрованныеМаркеры - Соответствие
// * ПаролиСертификатов - см. КриптографияБЭДКлиент.НовыеПаролиСертификатов
// * КонтекстДиагностики - см. ОбработкаНеисправностейБЭД.НовыйКонтекстДиагностики
// * ДанныеСертификатов - Соответствие
// * КлючиСинхронизации - см. СинхронизацияЭДОКлиентСервер.НовыеКлючиСинхронизации
// * ОповещениеОЗавершении - ОписаниеОповещения
Функция НовыйКонтекстПолученияПараметровАутентификации()
	
	Контекст = Новый Структура;
	Контекст.Вставить("КлючиСинхронизации", СинхронизацияЭДОКлиентСервер.НовыеКлючиСинхронизации());
	Контекст.Вставить("ДанныеСертификатов", Новый Соответствие);
	Контекст.Вставить("ЗашифрованныеМаркеры", Новый Соответствие);
	Контекст.Вставить("КонтекстДиагностики", Неопределено);
	Контекст.Вставить("ПаролиСертификатов", КриптографияБЭДКлиент.НовыеПаролиСертификатов());
	Контекст.Вставить("ОбработанныеУчетныеЗаписи", Новый Соответствие);
	Контекст.Вставить("ОповещениеОЗавершении", Неопределено);
	
	Возврат Контекст;
	
КонецФункции

Процедура ПолучитьПараметрыУведомленийПослеПолученияПараметровАутентификации(Результат, Контекст) Экспорт

	Контекст.КонтекстДиагностики = Результат.КонтекстДиагностики;
	Если Результат.КлючиСинхронизации.Количество() = 0 Тогда
		ВыполнитьОбработкуОповещения(Контекст.ОповещениеОЗавершении, РезультатПолученияПараметровУведомлений(Контекст));
		Возврат;
	КонецЕсли;
	
	Для Каждого КлючИЗначение Из Результат.КлючиСинхронизации Цикл
		КлючСинхронизации = КлючИЗначение.Значение;
	КонецЦикла;
	ПараметрыОжидания = ДлительныеОперацииКлиент.ПараметрыОжидания(Контекст.Форма);
	ПараметрыОжидания.ВыводитьОкноОжидания = Ложь;
	ПараметрыОжидания.ОповещениеПользователя.Показать = Ложь;
	ПараметрыОжидания.ВыводитьСообщения = Истина;

	ИдентификаторФормы = ?(Контекст.Форма = Неопределено, Новый УникальныйИдентификатор,
		Контекст.Форма.УникальныйИдентификатор);
	ДлительнаяОперация = СервисЭДОВызовСервера.НачатьПолучениеПараметровУведомлений(
			КлючСинхронизации, ИдентификаторФормы);
	
	Оповещение = Новый ОписаниеОповещения("ПослеПолученияПараметровУведомлений", ЭтотОбъект, Контекст);
	ДлительныеОперацииКлиент.ОжидатьЗавершение(ДлительнаяОперация, Оповещение, ПараметрыОжидания);

КонецПроцедуры

Процедура ПослеПолученияПараметровУведомлений(Результат, Контекст) Экспорт
	
	Если Результат.Статус = "Ошибка" Тогда
		Ошибка = ОбработкаНеисправностейБЭДКлиент.НоваяОшибка(
			НСтр("ru = 'Получение параметров уведомлений в сервисе электронного документооборота'"),
			ОбработкаНеисправностейБЭДКлиентСервер.ВидОшибкиНеизвестнаяОшибка(),
			Результат.КраткоеПредставлениеОшибки, Результат.КраткоеПредставлениеОшибки);
		ОбработкаНеисправностейБЭДКлиент.ДобавитьОшибку(Контекст.КонтекстДиагностики, Ошибка,
			ОбщегоНазначенияБЭДКлиентСервер.ПодсистемыБЭД().ОбменСКонтрагентами, Ложь);
	Иначе
		РезультатВыполнения = ПолучитьИзВременногоХранилища(Результат.АдресРезультата);
	КонецЕсли;
	
	Если РезультатВыполнения <> Неопределено Тогда
		Контекст.КонтекстДиагностики = РезультатВыполнения.КонтекстДиагностики;
		Контекст.ПараметрыУведомлений = РезультатВыполнения.ПараметрыУведомлений;
	КонецЕсли;
	
	ВыполнитьОбработкуОповещения(Контекст.ОповещениеОЗавершении, РезультатПолученияПараметровУведомлений(Контекст));
	
КонецПроцедуры

Процедура ОбновитьПараметрыУведомленийПослеПолученияПараметровАутентификации(Результат, Контекст) Экспорт
	
	Контекст.КонтекстДиагностики = Результат.КонтекстДиагностики;
	Если Результат.КлючиСинхронизации.Количество() = 0 Тогда
		ВыполнитьОбработкуОповещения(Контекст.ОповещениеОЗавершении, РезультатОбновленияПараметровУведомлений(Контекст));
		Возврат;
	КонецЕсли;
	
	Для Каждого КлючИЗначение Из Результат.КлючиСинхронизации Цикл
		КлючСинхронизации = КлючИЗначение.Значение;
	КонецЦикла;
	ПараметрыОжидания = ДлительныеОперацииКлиент.ПараметрыОжидания(Контекст.Форма);
	ПараметрыОжидания.ВыводитьОкноОжидания = Ложь;
	ПараметрыОжидания.ОповещениеПользователя.Показать = Ложь;
	ПараметрыОжидания.ВыводитьСообщения = Истина;
	
	ИдентификаторФормы = ?(Контекст.Форма = Неопределено, Новый УникальныйИдентификатор(),
		Контекст.Форма.УникальныйИдентификатор);
	ДлительнаяОперация = СервисЭДОВызовСервера.НачатьОбновлениеПараметровУведомлений(
		Контекст.ПараметрыУведомлений, КлючСинхронизации, ИдентификаторФормы);
	
	Оповещение = Новый ОписаниеОповещения("ПослеОбновленияПараметровУведомлений", ЭтотОбъект, Контекст);
	ДлительныеОперацииКлиент.ОжидатьЗавершение(ДлительнаяОперация, Оповещение, ПараметрыОжидания);
	
КонецПроцедуры

Процедура ПослеОбновленияПараметровУведомлений(Результат, Контекст) Экспорт
	
	Если Результат.Статус = "Ошибка" Тогда
		Ошибка = ОбработкаНеисправностейБЭДКлиент.НоваяОшибка(
			НСтр("ru = 'Обновление параметров уведомлений в сервисе электронного документооборота'"),
			ОбработкаНеисправностейБЭДКлиентСервер.ВидОшибкиНеизвестнаяОшибка(),
			Результат.КраткоеПредставлениеОшибки, Результат.КраткоеПредставлениеОшибки);
		ОбработкаНеисправностейБЭДКлиент.ДобавитьОшибку(Контекст.КонтекстДиагностики, Ошибка,
			ОбщегоНазначенияБЭДКлиентСервер.ПодсистемыБЭД().ОбменСКонтрагентами, Ложь);
	Иначе
		РезультатВыполнения = ПолучитьИзВременногоХранилища(Результат.АдресРезультата);
	КонецЕсли;
	
	Если РезультатВыполнения <> Неопределено Тогда
		Контекст.Успех = РезультатВыполнения.Результат;
		Контекст.КонтекстДиагностики = РезультатВыполнения.КонтекстДиагностики;
	КонецЕсли;
	
	ВыполнитьОбработкуОповещения(Контекст.ОповещениеОЗавершении, РезультатОбновленияПараметровУведомлений(Контекст));
	
КонецПроцедуры

Функция РезультатПолученияПараметровУведомлений(Контекст)
	
	Результат = Новый Структура;
	Результат.Вставить("КонтекстДиагностики", Контекст.КонтекстДиагностики);
	Результат.Вставить("ПараметрыУведомлений", Контекст.ПараметрыУведомлений);
	
	Возврат Результат;
	
КонецФункции

Функция РезультатОбновленияПараметровУведомлений(Контекст)
	
	Результат = Новый Структура;
	Результат.Вставить("Успех", Контекст.Успех);
	Результат.Вставить("КонтекстДиагностики", Контекст.КонтекстДиагностики);
	
	Возврат Результат;
	
КонецФункции

Процедура ОбработатьПредставлениеДанныхДляОператораЭДО(РезультатВыполнения, Контекст) Экспорт
	
	Если Контекст.Свойство("ДанныеПакета") Тогда
		ДанныеПакета = Контекст.ДанныеПакета;
	Иначе
		ДанныеПакета = Контекст;
	КонецЕсли;
	
	ОткрытьДанныеДляРегистрацииУОператораЭлектронногоДокументооборота(
		ДанныеПакета.РеквизитыПакета,
		ДанныеПакета.ДвоичныеДанныеДляОператораЭДО);
	
КонецПроцедуры

Процедура ОбработатьПредставлениеЗаявленияНаПодключениеКЭДО(РезультатВыполнения, Контекст) Экспорт
	
	ОткрытьПользовательскоеСоглашение1СЭДО();
	
КонецПроцедуры

#КОнецОбласти