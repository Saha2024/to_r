
#Область ПрограммныйИнтерфейс

#Область ИнтерактивныеДействия

// Функция-конструктор дополнительных параметров разбиения строки.
//
// Возвращаемое значение:
// 	Структура - с полями
//			*ИмяПоляКоличество - Строка - значение по умолчанию КоличествоУпаковок
//          *Заголовок - Строка - заголовок формы ввода числа, значение по умолчанию "Введите количество товара в новой строке"
//			*РазрешитьНулевоеКоличество - Булево - признак, что в исходной и конечной строке может быть 0, значение по умолчанию 0
//			*Количество - Неопределенно,Число - количество, которое будет отображено в форме редактирования числа. Если Неопределенно -
//					будет показано количество в текущей строке.
//					
Функция ПараметрыРазбиенияСтроки() Экспорт
	
	ПараметрыРазбиенияСтроки = Новый Структура;
	ПараметрыРазбиенияСтроки.Вставить("ИмяПоляКоличество", "КоличествоУпаковок");
	ПараметрыРазбиенияСтроки.Вставить("Заголовок", НСтр("ru = 'Введите количество товара в новой строке'"));
	ПараметрыРазбиенияСтроки.Вставить("РазрешитьНулевоеКоличество", Истина);
	ПараметрыРазбиенияСтроки.Вставить("Количество", Неопределено);
	
	Возврат ПараметрыРазбиенияСтроки;
	
КонецФункции

// Обработчик команды "Разбить строку"
//
//	Параметры:
//		ТЧ - ДанныеФормыКоллеция - табличная часть.
//		ЭлементФормы - ТаблицаФормы - элемент формы.
//		ПараметрыРазбиенияСтроки - Структура - см. ОбщегоНазначенияУТКлиент.ПараметрыРазбиенияСтроки.
//		ПотреблениеПоЗаказам - Булево - потребление по заказам.
//
Процедура РазбитьСтрокуТЧ(ТЧ, ЭлементФормы, ПараметрыРазбиенияСтроки = Неопределено, ПотреблениеПоЗаказам = Ложь) Экспорт
	
	Если ПараметрыРазбиенияСтроки = Неопределено Тогда
		ПараметрыОбработки = ПараметрыРазбиенияСтроки();
	Иначе
		ПараметрыОбработки = ПараметрыРазбиенияСтроки;
	КонецЕсли;
	
	ТекущаяСтрока	= ЭлементФормы.ТекущиеДанные;
	
	ЧислоВведено = Истина;
	
	Если ТекущаяСтрока = Неопределено Тогда
		ТекстСообщения = НСтр("ru = 'Для выполнения команды требуется выбрать строку табличной части.'");
		ПоказатьПредупреждение(,ТекстСообщения);
		Возврат;
	ИначеЕсли ТекущаяСтрока[ПараметрыОбработки.ИмяПоляКоличество] = 0
		И Не ПараметрыОбработки.РазрешитьНулевоеКоличество Тогда
		ТекстСообщения = НСтр("ru = 'Невозможно разбить строку с нулевым количеством.'");
		ПоказатьПредупреждение(,ТекстСообщения);
		Возврат;
	КонецЕсли;
	
	
	Если ПараметрыОбработки.Количество = Неопределено Тогда
		Количество = ТекущаяСтрока[ПараметрыОбработки.ИмяПоляКоличество];
	Иначе
		Количество = ПараметрыОбработки.Количество;
	КонецЕсли;
	
	СтруктураДопПараметров = Новый Структура("ПараметрыОбработки, ТЧ, ЭлементФормы, ПотреблениеПоЗаказам", ПараметрыОбработки, ТЧ, ЭлементФормы,ПотреблениеПоЗаказам);
	
	ОписаниеОповещения = Новый ОписаниеОповещения("РазбитьСтрокуТЧЗавершение", ОбщегоНазначенияУТКлиент, СтруктураДопПараметров);
	
	ПоказатьВводЧисла(ОписаниеОповещения, Количество, ПараметрыОбработки.Заголовок, 15, 3);
		
КонецПроцедуры

// Завершение предыдущей процедуры
//
// Параметры:
//  Число - Число - Число.
//  ДопПараметры - Структура - доп. параметры.
Процедура РазбитьСтрокуТЧЗавершение(Число, ДопПараметры) Экспорт
	
	// Правка ошибки при отмене операции
	Если Число = Неопределено Тогда
		Возврат;
	КонецЕсли;
	//
	
	ПараметрыОбработки   = ДопПараметры.ПараметрыОбработки;
	ТЧ                   = ДопПараметры.ТЧ;
	ЭлементФормы         = ДопПараметры.ЭлементФормы;
	ПотреблениеПоЗаказам = ДопПараметры.ПотреблениеПоЗаказам;
	ТекущаяСтрока	     = ЭлементФормы.ТекущиеДанные;
	Количество = Неопределено;	
	
	ОписаниеОповещения = Новый ОписаниеОповещения("РазбитьСтрокуТЧЗавершение", ОбщегоНазначенияУТКлиент, ДопПараметры);

	Если Не Число = Неопределено Тогда
		Количество = Число;
	КонецЕсли; 	
	
	Отказ = Ложь;
	
	Если Количество = 0
		И Не ПараметрыОбработки.РазрешитьНулевоеКоличество Тогда
		ТекстСообщения = НСтр("ru = 'Количество в новой строке не может быть равно нулю.'");
		
		ПоказатьПредупреждение(,ТекстСообщения);
		Отказ = Истина;	
	ИначеЕсли ТекущаяСтрока[ПараметрыОбработки.ИмяПоляКоличество] >= 0
		И Количество < 0 Тогда
		ТекстСообщения = НСтр("ru = 'Количество в новой строке не может быть отрицательным.'");
		ПоказатьПредупреждение(,ТекстСообщения);
		Отказ = Истина;	
	ИначеЕсли ТекущаяСтрока[ПараметрыОбработки.ИмяПоляКоличество] <= 0
		И Количество > 0 Тогда
		ТекстСообщения = НСтр("ru = 'Количество в новой строке не может быть положительным.'");
		ПоказатьПредупреждение(,ТекстСообщения);
		Отказ = Истина;	
	ИначеЕсли ТекущаяСтрока[ПараметрыОбработки.ИмяПоляКоличество] >= 0
		И Количество >  ТекущаяСтрока[ПараметрыОбработки.ИмяПоляКоличество] Тогда
		ТекстСообщения = НСтр("ru = 'Количество в новой строке не может быть больше количества в текущей.'");
		ПоказатьПредупреждение(,ТекстСообщения);
		Отказ = Истина;	
	ИначеЕсли ТекущаяСтрока[ПараметрыОбработки.ИмяПоляКоличество] <= 0
		И Количество <  ТекущаяСтрока[ПараметрыОбработки.ИмяПоляКоличество] Тогда
		ТекстСообщения = НСтр("ru = 'Количество в новой строке не может быть меньше количества в текущей.'");
		ПоказатьПредупреждение(,ТекстСообщения);
		Отказ = Истина;
	ИначеЕсли Количество =  ТекущаяСтрока[ПараметрыОбработки.ИмяПоляКоличество]
		И Не ПараметрыОбработки.РазрешитьНулевоеКоличество Тогда
		ТекстСообщения = НСтр("ru = 'Количество в новой строке должно отличаться от количества в текущей.'");
		ПоказатьПредупреждение(,ТекстСообщения);
		Отказ = Истина;
	КонецЕсли;
	
	Если Не Отказ Тогда
		
		ИндексТекущейСтроки 	 = ТЧ.Индекс(ТекущаяСтрока);
		НоваяСтрока 			 = ТЧ.Вставить(ИндексТекущейСтроки + 1);
		ЗаполнитьЗначенияСвойств(НоваяСтрока, ТекущаяСтрока);
		
		НоваяСтрока[ПараметрыОбработки.ИмяПоляКоличество]      = Количество;
		ТекущаяСтрока[ПараметрыОбработки.ИмяПоляКоличество]    = ТекущаяСтрока[ПараметрыОбработки.ИмяПоляКоличество]
		- НоваяСтрока[ПараметрыОбработки.ИмяПоляКоличество];
		
		ЭлементФормы.ТекущаяСтрока  = НоваяСтрока.ПолучитьИдентификатор();
		
		НоваяСтрока.КодСтроки = 0;
		НоваяСтрока.ЗаказНаВнутреннееПотребление = Неопределено;
		
		Если Не ПотреблениеПоЗаказам = Неопределено Тогда
			ОбновитьОтклоненияОтЗаказаВСтроке(НоваяСтрока, ПотреблениеПоЗаказам);
			Оповестить("ОбновитьРасхождениеЗаказа");
		КонецЕсли; 
		
		СтруктураДействий = Новый Структура;
		
		СтруктураДействий.Вставить("ПересчитатьКоличествоЕдиниц");	
		
		КэшированныеЗначения = Неопределено;
		
		ОбработкаТабличнойЧастиКлиент.ОбработатьСтрокуТЧ(ТекущаяСтрока, СтруктураДействий, КэшированныеЗначения);
		ОбработкаТабличнойЧастиКлиент.ОбработатьСтрокуТЧ(НоваяСтрока  , СтруктураДействий, КэшированныеЗначения);
		
	КонецЕсли;
КонецПроцедуры

// Процедура позволяет установить период через стандартный диалог выбора периода
//
// Параметры:
//  Объект                - Произвольный - Объект в котором устанавливается значения периода
//  ПараметрыПериода      - Структура - структура со свойствами "ДатаНачала", "ДатаОкончания" и в значениях имена полей
//                              объекта, для свойства "Вариант" - значение варианта стандартного периода.
//  ОповещениеПослеВыбора - ОписаниеОповещения - Описание оповещение которое выполняется после установки периода. 
//                              Может быть установлена пост-обрабокта в месте вызова после выбора периода.
// 
Процедура РедактироватьПериод(Объект, ПараметрыПериода = Неопределено, ОповещениеПослеВыбора = Неопределено) Экспорт
	
	Если ПараметрыПериода = Неопределено Тогда
		ПараметрыПериода = Новый Структура("ДатаНачала, ДатаОкончания", "ДатаНачала", "ДатаОкончания");
	КонецЕсли;
	
	Диалог = Новый ДиалогРедактированияСтандартногоПериода();
	Если ПараметрыПериода.Свойство("ДатаНачала") Тогда
		Диалог.Период.ДатаНачала = Объект[ПараметрыПериода.ДатаНачала];
	КонецЕсли; 
	Если ПараметрыПериода.Свойство("ДатаОкончания") Тогда
		Диалог.Период.ДатаОкончания = Объект[ПараметрыПериода.ДатаОкончания];
	КонецЕсли; 
	Если ПараметрыПериода.Свойство("Вариант") Тогда
		Диалог.Период.Вариант = ПараметрыПериода.Вариант;
	КонецЕсли; 
	
	ДополнительныеПараметры = Новый Структура;
	ДополнительныеПараметры.Вставить("Объект",           Объект);
	ДополнительныеПараметры.Вставить("ПараметрыПериода", ПараметрыПериода);
	Если ОповещениеПослеВыбора <> Неопределено Тогда
		ДополнительныеПараметры.Вставить("ОповещениеПослеВыбора", ОповещениеПослеВыбора);
	КонецЕсли; 
	
	Оповещение = Новый ОписаниеОповещения(
		"РедактироватьПериодЗавершение", 
		ЭтотОбъект, 
		ДополнительныеПараметры);
	
	Диалог.Показать(Оповещение);

КонецПроцедуры

// Позволяет выбрать период из выпадающего списка
//
// Параметры:
//		Форма				- Форма - форма, в которой выполняется выбор периода.
//		Элемент				- ЭлементФормы - элемент, рядом с которым нужно открыть выпадающий список.
//		ВидПериода			- ПеречислениеСсылка.ДоступныеПериодыОтчета - вид выбираемого периода.
//		ОбработчикВыбора	- ОписаниеОповещения - обработчик, расположенный в форме, который будет вызван после выбора.
//		НачалоПериода		- Дата - начало периода.
//
Процедура ВыбратьПериодИзСписка(Форма, Элемент, ВидПериода, НачалоПериода, ОбработчикВыбора) Экспорт
	
	Если НачалоПериода = '00010101' Тогда
		НачалоПериода = ОбщегоНазначенияУТКлиентСервер.НачалоПериодаОтчета(ВидПериода, ТекущаяДата());
	КонецЕсли;
	
	ВыбратьПериодОтчета(Форма, Элемент, ВидПериода, НачалоПериода, ОбработчикВыбора);
	
КонецПроцедуры

Функция ВыбратьПериодОтчетаПоПараметрам(Форма, Элемент, СтандартнаяОбработка, ВидПериода, НачалоПериода)
	
	Список = ОбщегоНазначенияУТКлиентСервер.ПолучитьСписокПериодов(НачалоПериода, ВидПериода);
	Если Список.Количество() = 0 Тогда
		СтандартнаяОбработка = Ложь;
		Возврат Неопределено;
	КонецЕсли;
	
	ЭлементСписка = Список.НайтиПоЗначению(НачалоПериода);
	ВыбранныйПериод = Форма.ВыбратьИзСписка(Список, Элемент, ЭлементСписка);
	
	Если ВыбранныйПериод = Неопределено тогда
		Возврат Неопределено;
	КонецЕсли;
	
	Индекс = Список.Индекс(ВыбранныйПериод);
	Если Индекс = 0 ИЛИ Индекс = Список.Количество() - 1 тогда
		ВыбранныйПериод = ВыбратьПериодОтчетаПоПараметрам(Форма, Элемент, СтандартнаяОбработка, ВидПериода, ВыбранныйПериод.Значение);
	КонецЕсли;
	
	Возврат ВыбранныйПериод;
	
КонецФункции

// Процедура - Период начало выбора
//
// Параметры:
//  Форма				 - ФормаКлиентскогоПриложения - форма.
//  Элемент				 - ДекорацияФормы, ГруппаФормы, КнопкаФормы, ТаблицаФормы, ПолеФормы - элемент формы.
//  СтандартнаяОбработка - Булево - стандартная обработка.
//  ВидПериода			 - Перечисление.ДоступныеПериодыОтчета - вид периода.
//  Период				 - Период - Период.
//  НачалоПериода		 - Дата - Начало периода.
//  КонецПериода		 - Дата - КонецПериода.
Процедура ПериодНачалоВыбора(Форма, Элемент, СтандартнаяОбработка, ВидПериода, Период, НачалоПериода, КонецПериода) Экспорт

	Если НачалоПериода = '00010101' Тогда
		НачалоПериода = ОбщегоНазначенияУТКлиентСервер.НачалоПериодаОтчета(ВидПериода, ТекущаяДата());
	КонецЕсли;
	ВыбранныйПериод      = ВыбратьПериодОтчетаПоПараметрам(Форма, Элемент, СтандартнаяОбработка, ВидПериода, НачалоПериода);
	СтандартнаяОбработка = Ложь;
	Если ВыбранныйПериод = Неопределено Тогда
		Возврат;
	КонецЕсли;
	Период = ВыбранныйПериод.Представление;
	
	НачалоПериода = ВыбранныйПериод.Значение;
	КонецПериода  = ОбщегоНазначенияУТКлиентСервер.КонецПериодаОтчета(ВидПериода, ВыбранныйПериод.Значение);
	
КонецПроцедуры

#КонецОбласти

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

Процедура ОбновитьОтклоненияОтЗаказаВСтроке(ТекущаяСтрока, ПотреблениеПоЗаказам) Экспорт
	
	Если ПотреблениеПоЗаказам Тогда
		Если НЕ ЗначениеЗаполнено(ТекущаяСтрока.ЗаказНаВнутреннееПотребление) Тогда
			ТекущаяСтрока.РасхождениеЗаказ = 1;
		Иначе
			ТекущаяСтрока.РасхождениеЗаказ = 0;
		КонецЕсли;
	КонецЕсли;
КонецПроцедуры 

#Область ВыборПериодаИзСписка

Процедура ВыбратьПериодОтчета(Форма, Элемент, ВидПериода, НачалоПериода, ОбработчикВыбора)
	
	Список = ОбщегоНазначенияУТКлиентСервер.ПолучитьСписокПериодов(НачалоПериода, ВидПериода);
	Если Список.Количество() = 0 Тогда
		Возврат;
	КонецЕсли;
	
	ЭлементСписка = Список.НайтиПоЗначению(НачалоПериода);
	
	ДополнительныеПараметры = Новый Структура;
	ДополнительныеПараметры.Вставить("Форма", Форма);
	ДополнительныеПараметры.Вставить("Элемент", Элемент);
	ДополнительныеПараметры.Вставить("ВидПериода", ВидПериода);
	ДополнительныеПараметры.Вставить("НачалоПериода", НачалоПериода);
	ДополнительныеПараметры.Вставить("Список", Список);
	ДополнительныеПараметры.Вставить("ОбработчикВыбора", ОбработчикВыбора);
	ОписаниеОповещения = Новый ОписаниеОповещения("ВыбратьПериодОтчетаЗавершение", ЭтотОбъект, ДополнительныеПараметры);
	
	Форма.ПоказатьВыборИзСписка(ОписаниеОповещения, Список, Элемент, ЭлементСписка);
	
КонецПроцедуры

Процедура ВыбратьПериодОтчетаЗавершение(ВыбранныйПериод, ДополнительныеПараметры) Экспорт

	Если ВыбранныйПериод = Неопределено тогда
		Возврат;
	КонецЕсли;
	
	Индекс = ДополнительныеПараметры.Список.Индекс(ВыбранныйПериод);
	Если Индекс = 0 ИЛИ Индекс = ДополнительныеПараметры.Список.Количество() - 1 тогда
		ВыбратьПериодОтчета(ДополнительныеПараметры.Форма, 
							ДополнительныеПараметры.Элемент, 
							ДополнительныеПараметры.ВидПериода, 
							ВыбранныйПериод.Значение, 
							ДополнительныеПараметры.ОбработчикВыбора);
		Возврат;
	КонецЕсли;
	
	ВыполнитьОбработкуОповещения(ДополнительныеПараметры.ОбработчикВыбора, ВыбранныйПериод);
	
КонецПроцедуры

// Процедура завершения для РедактироватьПериод()
// см. подробней для процедуры РедактироватьПериод().
Процедура РедактироватьПериодЗавершение(Период, ДополнительныеПараметры) Экспорт 

	ПараметрыПериода = ДополнительныеПараметры.ПараметрыПериода;
	Объект           = ДополнительныеПараметры.Объект;
	Если Период <> Неопределено Тогда
		Если ПараметрыПериода.Свойство("ДатаНачала") Тогда
			Объект[ПараметрыПериода.ДатаНачала]= Период.ДатаНачала;
		КонецЕсли; 
		Если ПараметрыПериода.Свойство("ДатаОкончания") Тогда
			Объект[ПараметрыПериода.ДатаОкончания]= Период.ДатаОкончания;
		КонецЕсли; 
		Если ПараметрыПериода.Свойство("Вариант") Тогда
			Объект[ПараметрыПериода.Вариант]= Период.Вариант;
		КонецЕсли;
	КонецЕсли;
	
	Если ДополнительныеПараметры.Свойство("ОповещениеПослеВыбора") Тогда
		ВыполнитьОбработкуОповещения(ДополнительныеПараметры.ОповещениеПослеВыбора, Период);
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти

#КонецОбласти
