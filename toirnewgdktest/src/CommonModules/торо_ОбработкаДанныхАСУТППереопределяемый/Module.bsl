
#Область ПрограммныйИнтерфейс

// Функция определяет обработчики создания документов по данным АСУ ТП,
// которые требуется переопределять. По умолчанию все нестандартные обработчики отключены.
// 
// Возвращаемое значение:
//		Соответствие - соответствие с полями:
//			*Ключ - ПеречислениеСсылка.торо_ТипыСигналовАСУТП - вид документа.
//			*Значение - Булево - Истина, если нужен нестандартный обработчик.
//
Функция ПолучитьСписокПереопределяемыхФункцийОбработчиков() Экспорт 
	
	Обработчики = Новый Соответствие;
	Обработчики.Вставить(Перечисления.торо_ТипыСигналовАСУТП.Состояние, Ложь);
	Обработчики.Вставить(Перечисления.торо_ТипыСигналовАСУТП.Показатель, Ложь);
	Обработчики.Вставить(Перечисления.торо_ТипыСигналовАСУТП.Наработка, Ложь);
	Обработчики.Вставить(Перечисления.торо_ТипыСигналовАСУТП.Дефект, Ложь);
	Обработчики.Вставить(Перечисления.торо_ТипыСигналовАСУТП.KPI, Ложь);
	
	Возврат Обработчики;
	
КонецФункции

// В процедуре описывается нестандартный алгоритм создания документов "Состояния объектов ремонта".
// Процедура будет вызвана только, если в процедуре ПолучитьСписокПереопределяемыхФункцийОбработчиков
// для ключа Перечисления.торо_ТипыСигналовАСУТП.Состояние установлено значение Истина.
//
// Параметры:
//  Выборка						 - СтрокаТаблицыЗначений - строка с данными для создания документа.
//  масСозданныхДокументов	 - Массив - массив, в который следует добавить ссылки на созданные документы.
//  сооДокументыИИсхДанные	 - Соответствие - соответствие, в которое следует добавить ссылки на созданные документы:
//										*Ключ - ДокументСсылка - ссылка на созданный документ.
//										*Значение - СтрокаТаблицыЗначений - значение параметра Выборка.
//  ОсновнаяИерархия			 - СправочникСсылка.торо_СтруктурыОР - структура иерархии объектов ремонта.
//
Процедура СостояниеОР(Выборка, масСозданныхДокументов, сооДокументыИИсхДанные, ОсновнаяИерархия) Экспорт
	
КонецПроцедуры

// В процедуре описывается нестандартный алгоритм создания документов "Учет контролируемых показателей".
// Процедура будет вызвана только, если в процедуре ПолучитьСписокПереопределяемыхФункцийОбработчиков
// для ключа Перечисления.торо_ТипыСигналовАСУТП.Показатель установлено значение Истина.
//
// Параметры:
//  Выборка						 - СтрокаТаблицыЗначений - строка с данными для создания документа.
//  масСозданныхДокументов	 - Массив - массив, в который следует добавить ссылки на созданные документы.
//  сооДокументыИИсхДанные	 - Соответствие - соответствие, в которое следует добавить ссылки на созданные документы:
//										*Ключ - ДокументСсылка - ссылка на созданный документ.
//										*Значение - СтрокаТаблицыЗначений - значение параметра Выборка.
//  ОсновнаяИерархия			 - СправочникСсылка.торо_СтруктурыОР - структура иерархии объектов ремонта.
//
Процедура Показатель(Выборка, масСозданныхДокументов, сооДокументыИИсхДанные, ОсновнаяИерархия) Экспорт
	
КонецПроцедуры

// В процедуре описывается нестандартный алгоритм создания документов "Учет наработки оборудования".
// Процедура будет вызвана только, если в процедуре ПолучитьСписокПереопределяемыхФункцийОбработчиков
// для ключа Перечисления.торо_ТипыСигналовАСУТП.Наработка установлено значение Истина.
//
// Параметры:
//  Выборка						 - СтрокаТаблицыЗначений - строка с данными для создания документа.
//  масСозданныхДокументов	 - Массив - массив, в который следует добавить ссылки на созданные документы.
//  сооДокументыИИсхДанные	 - Соответствие - соответствие, в которое следует добавить ссылки на созданные документы:
//										*Ключ - ДокументСсылка - ссылка на созданный документ.
//										*Значение - СтрокаТаблицыЗначений - значение параметра Выборка.
//  ОсновнаяИерархия			 - СправочникСсылка.торо_СтруктурыОР - структура иерархии объектов ремонта.
//
Процедура Наработка(Выборка, масСозданныхДокументов, сооДокументыИИсхДанные, ОсновнаяИерархия) Экспорт
	
КонецПроцедуры

// В процедуре описывается нестандартный алгоритм создания документов "Выявленные дефекты".
// Процедура будет вызвана только, если в процедуре ПолучитьСписокПереопределяемыхФункцийОбработчиков
// для ключа Перечисления.торо_ТипыСигналовАСУТП.Дефект установлено значение Истина.
//
// Параметры:
//  Выборка						 - СтрокаТаблицыЗначений - строка с данными для создания документа.
//  масСозданныхДокументов	 - Массив - массив, в который следует добавить ссылки на созданные документы.
//  сооДокументыИИсхДанные	 - Соответствие - соответствие, в которое следует добавить ссылки на созданные документы:
//										*Ключ - ДокументСсылка - ссылка на созданный документ.
//										*Значение - СтрокаТаблицыЗначений - значение параметра Выборка.
//  ОсновнаяИерархия			 - СправочникСсылка.торо_СтруктурыОР - структура иерархии объектов ремонта.
//
Процедура Дефект(Выборка, масСозданныхДокументов, сооДокументыИИсхДанные, ОсновнаяИерархия) Экспорт
	
КонецПроцедуры

// В процедуре описывается нестандартный алгоритм создания документов "Регистрация значений внешних показателей".
// Процедура будет вызвана только, если в процедуре ПолучитьСписокПереопределяемыхФункцийОбработчиков
// для ключа Перечисления.торо_ТипыСигналовАСУТП.KPI установлено значение Истина.
//
// Параметры:
//  Выборка						 - СтрокаТаблицыЗначений - строка с данными для создания документа.
//  масСозданныхДокументов	 - Массив - массив, в который следует добавить ссылки на созданные документы.
//  сооДокументыИИсхДанные	 - Соответствие - соответствие, в которое следует добавить ссылки на созданные документы:
//										*Ключ - ДокументСсылка - ссылка на созданный документ.
//										*Значение - СтрокаТаблицыЗначений - значение параметра Выборка.
//  ОсновнаяИерархия			 - СправочникСсылка.торо_СтруктурыОР - структура иерархии объектов ремонта.
//
Процедура KPI(Выборка, масСозданныхДокументов, сооДокументыИИсхДанные, ОсновнаяИерархия) Экспорт
	
КонецПроцедуры

#КонецОбласти
