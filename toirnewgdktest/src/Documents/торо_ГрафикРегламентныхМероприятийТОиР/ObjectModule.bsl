#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда
	
////////////////////////////////////////////////////////////////////////////////
// СЛУЖЕБНЫЕ ПЕРЕМЕННЫЕ

перем СтруктураДанных Экспорт; // Содержит структуру данных о статусе документа (Начальный, СогласованиеЗавершено, ПроводитьДокумент, ШаблонСообщенияЭлектроннойПочты) 
										 // из регистра сведений торо_МатрицаПереходаСтатусовДокументов. 
Перем БезусловнаяЗапись Экспорт; // Отключает проверки при записи документа

#Область ОбработчикиСобытий

Процедура ОбработкаЗаполнения(Основание)
	
	Ответственный = ОбщегоНазначения.ХранилищеОбщихНастроекЗагрузить("НастройкиТОиР", "ОсновнойОтветственный");
	Если НЕ ЗначениеЗаполнено(Ответственный) тогда
		Ответственный = Пользователи.ТекущийПользователь();
	КонецЕсли;	
	
	Если ТипЗнч(Основание) = Тип("ДокументСсылка.торо_ГрафикРегламентныхМероприятийТОиР") Тогда
		
		ВидОперации = Перечисления.торо_ВидыОперацийПланаГрафикаППР.Корректировка;
		Организация = Основание.Организация;
		Подразделение = Основание.Подразделение;
		ДокументОснование = Основание.Ссылка;
		
		ДатаПланирования = Основание.ДатаПланирования;
		КоличествоПериодов = Основание.КоличествоПериодов;
		ПериодичностьДетализации = Основание.ПериодичностьДетализации;
			
		Если ЗначениеЗаполнено(ПериодичностьДетализации) Тогда		
			мДатаОкончанияПланирования = ОбщегоНазначенияУТКлиентСервер.РассчитатьДатуОкончанияПериода(ДатаПланирования, ПериодичностьДетализации, КоличествоПериодов);
		Иначе
			ТекстСообщения = НСтр("ru = 'Не выбран период планирования!'");
			ОбщегоНазначения.СообщитьПользователю(ТекстСообщения);
		КонецЕсли;
		
		Если мДатаОкончанияПланирования <> Неопределено Тогда
			Документы.торо_ГрафикРегламентныхМероприятийТОиР.ЗаполнитьМаршрутыИОбъектыМаршрутов(Маршруты, Основание);
			Документы.торо_ГрафикРегламентныхМероприятийТОиР.ЗаполнитьПланРемонтов(ПланРемонтов, Основание, ДатаПланирования, мДатаОкончанияПланирования, Маршруты.Выгрузить());
		КонецЕсли;
		
	КонецЕсли;
	
КонецПроцедуры

Процедура ОбработкаПроверкиЗаполнения(Отказ, ПроверяемыеРеквизиты)
	
	МассивНепроверяемыхРеквизитов = Новый массив;

	Если ВидОперации <> Перечисления.торо_ВидыОперацийПланаГрафикаППР.Корректировка Тогда
		МассивНепроверяемыхРеквизитов.Добавить("ДокументОснование");	
	КонецЕсли;
	
	Если ПланРемонтов.Количество() = 0 Тогда
		ТекстСообщения = НСтр("ru = 'План график не содержит данных.'");
		ОбщегоНазначения.СообщитьПользователю(ТекстСообщения,,,,Отказ);
	КонецЕсли;	
		
	ПроверитьЗаполнениеТабличнойЧастиПланРемонтов(Отказ);
	
	ОбщегоНазначения.УдалитьНепроверяемыеРеквизитыИзМассива(ПроверяемыеРеквизиты, МассивНепроверяемыхРеквизитов);
	
	Если ВидОперации = Перечисления.торо_ВидыОперацийПланаГрафикаППР.ПланРабот Тогда
		ПроверитьПериодыПланированияРемонтныхРабот(Отказ);
	КонецЕсли;
	
КонецПроцедуры

Процедура ПередЗаписью(Отказ, РежимЗаписи, РежимПроведения)
	// Получение удаленных ремонтов при перепроведении документа. Используются при формировании записей
	// регистров торо_ПлановыеИсполнителиРемонтов, торо_ОбщиеДанныеПоРемонтам, торо_СвернутыеФактическиеДатыРемонтов.
	Если Проведен И РежимЗаписи = РежимЗаписиДокумента.Проведение Тогда
		МассивIDДокумента = ПланРемонтов.ВыгрузитьКолонку("ID");
		МассивУдаленныхID = торо_Ремонты.ПолучитьIDУдаленныхРемонтовДокумента(МассивIDДокумента, Ссылка);
		ДополнительныеСвойства.Вставить("МассивУдаленныхID", МассивУдаленныхID);
	Иначе
		ДополнительныеСвойства.Вставить("МассивУдаленныхID", Новый Массив());
	КонецЕсли;
	
	Если ЭтоНовый() Тогда
		Автор = Пользователи.ТекущийПользователь();
	КонецЕсли;
КонецПроцедуры

Процедура ОбработкаПроведения(Отказ, РежимПроведения)
	
	Если ЗначениеЗаполнено(ДокументОснование) Тогда
		Если Дата <= ОбщегоНазначения.ЗначениеРеквизитаОбъекта(ДокументОснование, "Дата") Тогда
			ШаблонСообщения = НСтр("ru = 'Дата документа должна быть больше даты документа основания ''%1''!'");
			ТекстСообщения = СтрШаблон(ШаблонСообщения, Строка(ДокументОснование));
			ОбщегоНазначения.СообщитьПользователю(ТекстСообщения,,,,Отказ);
			Возврат;
		ИначеЕсли Не ОбщегоНазначения.ЗначениеРеквизитаОбъекта(ДокументОснование, "Проведен") Тогда
			ШаблонСообщения = НСтр("ru = 'Есть непроведенный документ основание: %1'");
			ТекстСообщения = СтрШаблон(ШаблонСообщения, Строка(ДокументОснование));
			ОбщегоНазначения.СообщитьПользователю(ТекстСообщения,,,,Отказ);
			Возврат;			
		КонецЕсли;
	КонецЕсли;	
	
	ТЗ = Новый ТаблицаЗначений;
	ТЗ = ПланРемонтов.Выгрузить();
	ТЗ.Свернуть("СписокОбъектовРемонта");
	
	Запрос = Новый Запрос;
	Запрос.Текст = "ВЫБРАТЬ
	               |	ТабДоступность.ID КАК ID,
	               |	ТабДоступность.ДоступенДляРедактирования КАК ДоступенДляРедактирования
	               |ПОМЕСТИТЬ ТабДоступность
	               |ИЗ
	               |	&ТабДоступность КАК ТабДоступность
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ РАЗРЕШЕННЫЕ
	               |	торо_ПланГрафикРемонтаПланРемонтов.НомерСтроки КАК НомерСтроки,
	               |	торо_ПланГрафикРемонтаПланРемонтов.ID КАК ID,
	               |	торо_ПланГрафикРемонтаПланРемонтов.ВидРемонтныхРабот КАК ВидРемонтныхРабот,
	               |	торо_ПланГрафикРемонтаПланРемонтов.ДатаКон КАК ДатаКон,
	               |	торо_ПланГрафикРемонтаПланРемонтов.ДатаНач КАК ДатаНач,
	               |	торо_ПланГрафикРемонтаПланРемонтов.Исполнитель КАК Исполнитель,
	               |	торо_ПланГрафикРемонтаПланРемонтов.СписокОбъектовРемонта КАК СписокОбъектовРемонта,
	               |	торо_ПланГрафикРемонтаПланРемонтов.СпособВыполнения КАК СпособВыполнения,
	               |	торо_ПланГрафикРемонтаПланРемонтов.Отменен КАК Отменен,
	               |	торо_ПланГрафикРемонтаПланРемонтов.Склад КАК Склад,
	               |	торо_ПланГрафикРемонтаПланРемонтов.Замещен КАК Замещен,
	               |	торо_ПланГрафикРемонтаПланРемонтов.ID_базы_расчета КАК ID_базы_расчета,
	               |	торо_ПланГрафикРемонтаПланРемонтов.ID_замещающего КАК ID_замещающего,
	               |	торо_ПланГрафикРемонтаПланРемонтов.ДатаНачСт КАК ДатаНачСт,
	               |	торо_ПланГрафикРемонтаПланРемонтов.ДатаКонСт КАК ДатаКонСт,
	               |	торо_ПланГрафикРемонтаПланРемонтов.ЗамещенСт КАК ЗамещенСт,
	               |	торо_ПланГрафикРемонтаПланРемонтов.СуммаРемонта КАК СуммаРемонта,
	               |	торо_ПланГрафикРемонтаПланРемонтов.СрокПоНормативу КАК СрокПоНормативу
	               |ПОМЕСТИТЬ ПланРемонтов
	               |ИЗ
	               |	Документ.торо_ГрафикРегламентныхМероприятийТОиР.ПланРемонтов КАК торо_ПланГрафикРемонтаПланРемонтов
	               |ГДЕ
	               |	торо_ПланГрафикРемонтаПланРемонтов.Ссылка = &Ссылка
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ РАЗЛИЧНЫЕ
	               |	торо_ПланГрафикРемонтаПланРемонтов.ID КАК ID,
	               |	торо_ПланГрафикРемонтаПланРемонтов.ВидРемонтныхРабот КАК ВидРемонтныхРабот,
	               |	торо_ПланГрафикРемонтаПланРемонтов.ДатаКон КАК ДатаКон,
	               |	торо_ПланГрафикРемонтаПланРемонтов.ДатаНач КАК ДатаНач,
	               |	торо_ПланГрафикРемонтаПланРемонтов.Исполнитель КАК Исполнитель,
	               |	торо_ПланГрафикРемонтаПланРемонтов.СписокОбъектовРемонта КАК СписокОбъектовРемонта,
	               |	торо_ПланГрафикРемонтаПланРемонтов.СпособВыполнения КАК СпособВыполнения,
	               |	торо_ПланГрафикРемонтаПланРемонтов.Отменен КАК Отменен,
	               |	торо_ПланГрафикРемонтаПланРемонтов.Склад КАК Склад,
	               |	торо_ПланГрафикРемонтаПланРемонтов.Замещен КАК Замещен,
	               |	торо_ПланГрафикРемонтаПланРемонтов.ID_базы_расчета КАК ID_базы_расчета,
	               |	торо_ПланГрафикРемонтаПланРемонтов.ID_замещающего КАК ID_замещающего,
	               |	торо_ПланГрафикРемонтаПланРемонтов.ДатаНачСт КАК ДатаНачСт,
	               |	торо_ПланГрафикРемонтаПланРемонтов.ДатаКонСт КАК ДатаКонСт,
	               |	торо_ПланГрафикРемонтаПланРемонтов.ЗамещенСт КАК ЗамещенСт,
	               |	торо_ПланГрафикРемонтаПланРемонтов.СуммаРемонта КАК СуммаРемонта,
	               |	торо_ПланГрафикРемонтаПланРемонтов.СрокПоНормативу КАК СрокПоНормативу
	               |ИЗ
	               |	ПланРемонтов КАК торо_ПланГрафикРемонтаПланРемонтов
	               |		ВНУТРЕННЕЕ СОЕДИНЕНИЕ ТабДоступность КАК ТабДоступность
	               |		ПО торо_ПланГрафикРемонтаПланРемонтов.ID = ТабДоступность.ID
	               |ГДЕ
	               |	ТабДоступность.ДоступенДляРедактирования";
	
	Запрос.УстановитьПараметр("Ссылка", Ссылка);
	Запрос.УстановитьПараметр("ТабДоступность", Документы.торо_ГрафикРегламентныхМероприятийТОиР.ОбновитьДоступностьДляРедактирования(Ссылка,ВидОперации,ПланРемонтов.Выгрузить()));
	ПланРемонтовДвижения = Запрос.Выполнить().Выгрузить();
	
	Если ЗначениеЗаполнено(ПериодичностьДетализации) Тогда		
		ДатаОкончанияПланирования = ОбщегоНазначенияУТКлиентСервер.РассчитатьДатуОкончанияПериода(ДатаПланирования, ПериодичностьДетализации, КоличествоПериодов);
	Иначе
		ТекстСообщения = НСтр("ru = 'Не выбран период планирования!'");
		ОбщегоНазначения.СообщитьПользователю(ТекстСообщения);
	КонецЕсли;
	
	ДвиженияПоРегистру_торо_ПериодыПланированияППР(РежимПроведения, ДатаОкончанияПланирования, Отказ, ПланРемонтовДвижения);
	
	ДвиженияПоРегистру_торо_ПлановыеМероприятия(РежимПроведения, Отказ, ПланРемонтовДвижения);
	
	// Движение по регистру "торо_ПлановыеРемонтныеРаботы".
	ДвиженияПоРегистру_торо_ЗапланированныеМероприятияОбъектов(РежимПроведения, Отказ, ПланРемонтовДвижения);
	
	ДвиженияПоРегистру_торо_СоответствиеКорректировокДокументамППР(РежимПроведения, Отказ);
	
	ДвиженияПоРегистру_торо_ВсеДокументыПоРемонтам();
	
	торо_РаботаСоСтатусамиДокументовСервер.УстановитьСтатусРемонтовПриПроведении(ЭтотОбъект);
	торо_МТОСервер.ПоместитьРемонтыВРегистрДляПроверкиАктуальности(Ссылка);
	
	торо_Ремонты.ОбновитьЗаписиНезависимыхРегистровПоРемонтам(ЭтотОбъект, РежимЗаписиДокумента.Проведение);
	
КонецПроцедуры

Процедура ОбработкаУдаленияПроведения(Отказ)
	
	торо_МТОСервер.ПоместитьРемонтыВРегистрДляПроверкиАктуальности(Ссылка,, Истина);
		
	Если Не Отказ Тогда
		торо_Ремонты.ОбновитьЗаписиНезависимыхРегистровПоРемонтам(ЭтотОбъект, РежимЗаписиДокумента.ОтменаПроведения);
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции
Процедура ДвиженияПоРегистру_торо_СоответствиеКорректировокДокументамППР(РежимПроведения, Отказ)
	Если ВидОперации = Перечисления.торо_ВидыОперацийПланаГрафикаППР.Корректировка Тогда
		Если ЗначениеЗаполнено(ДокументОснование) Тогда
			
			Движения.торо_СоответствиеКорректировокДокументамППР.Очистить();
			Движения.торо_СоответствиеКорректировокДокументамППР.Записывать = Истина;
			
			Если ДокументОснование.ВидОперации = Перечисления.торо_ВидыОперацийПланаГрафикаППР.Корректировка Тогда
				ИсходныйППР = Неопределено;
				Запрос = Новый Запрос;
				Запрос.Текст = "ВЫБРАТЬ
				|	торо_СоответствиеКорректировокДокументамППР.Корректируемый
				|ИЗ
				|	РегистрСведений.торо_СоответствиеКорректировокДокументамППР КАК торо_СоответствиеКорректировокДокументамППР
				|ГДЕ
				|	торо_СоответствиеКорректировокДокументамППР.Корректировка = &Корректировка";
				Запрос.УстановитьПараметр("Корректировка",ДокументОснование);
				
				НаборИсходных = Запрос.Выполнить().Выгрузить();
				СтрНабора = НаборИсходных.Добавить();
				СтрНабора.Корректируемый = ДокументОснование;
			Иначе
				ИсходныйППР = ДокументОснование;					
			КонецЕсли;
			
			Если ИсходныйППР <> Неопределено Тогда
				Движение = Движения.торо_СоответствиеКорректировокДокументамППР.Добавить();
				Движение.Корректировка = Ссылка;
				Движение.Корректируемый = ИсходныйППР;
			Иначе
				Для Каждого ЗаписьИсходный Из НаборИсходных Цикл
					Движение = Движения.торо_СоответствиеКорректировокДокументамППР.Добавить();
					Движение.Корректировка = Ссылка;
					Движение.Корректируемый = ЗаписьИсходный.Корректируемый;
				КонецЦикла;
			КонецЕсли;
		КонецЕсли;
		
	КонецЕсли;
	
КонецПроцедуры

Процедура ДвиженияПоРегистру_торо_ПлановыеМероприятия(РежимПроведения, Отказ, ПланРемонтовДвижения)
	
	Если ПланРемонтовДвижения.Количество() = 0 Тогда
		Возврат;
	КонецЕсли;
	
	НаборЗаписей = РегистрыСведений.торо_ПлановыеМероприятия.СоздатьНаборЗаписей();
	НаборЗаписей.Отбор.Регистратор.Установить(Ссылка);
	НаборЗаписей.Прочитать();
	СтарыеДвижения = НаборЗаписей.Выгрузить();
	
	Движения.торо_ПлановыеМероприятия.Очистить();	
	Движения.торо_ПлановыеМероприятия.Записывать = Истина;
	
	Если Не Отказ Тогда
		
		Для Каждого ИсходнаяСтрока Из СтарыеДвижения Цикл
			СтрДостДляРед = ПланРемонтовДвижения.НайтиСтроки(Новый Структура("ID", ИсходнаяСтрока.ID));
			СтрЕстьВДок = ПланРемонтов.НайтиСтроки(Новый Структура("ID", ИсходнаяСтрока.ID));
			Если СтрДостДляРед.Количество() = 0 И СтрЕстьВДок.Количество() > 0 Тогда
				Движение = Движения.торо_ПлановыеМероприятия.Добавить();
				ЗаполнитьЗначенияСвойств(Движение,ИсходнаяСтрока);
			КонецЕсли;
		КонецЦикла;
		
		Для каждого ИсходнаяСтрока Из ПланРемонтовДвижения Цикл
			Движение = Движения.торо_ПлановыеМероприятия.Добавить();
			ЗаполнитьЗначенияСвойств(Движение, ИсходнаяСтрока);
			Движение.Период = Дата;
			Движение.Организация = Организация;
			Движение.Подразделение = Подразделение;
			Движение.СписокОбъектов = ИсходнаяСтрока.СписокОбъектовРемонта;
			Движение.ВидМероприятия = ИсходнаяСтрока.ВидРемонтныхРабот;
			Движение.СтоимостьРемонта = ИсходнаяСтрока.СуммаРемонта;
			Движение.ДатаНачала = ИсходнаяСтрока.ДатаНач;
			Движение.ДатаОкончания = ИсходнаяСтрока.ДатаКон;
		КонецЦикла;
		
	КонецЕсли;
		
КонецПроцедуры

Процедура ДвиженияПоРегистру_торо_ЗапланированныеМероприятияОбъектов(РежимПроведения, Отказ, ПланРемонтовДвижения)
	
	Запрос = Новый Запрос;
	Запрос.Текст =
	"ВЫБРАТЬ
	|	ПланРемонтныхРабот.ID КАК ID,
	|	ПланРемонтныхРабот.СписокОбъектовРемонта КАК СписокОбъектовРемонта
	|ПОМЕСТИТЬ ДанныеИзПланаРемонтов
	|ИЗ
	|	&ПланРемонтныхРабот КАК ПланРемонтныхРабот
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	торо_МаршрутыРегламентныхМероприятий.ОбъектРемонта КАК ОбъектРемонта,
	|	ДанныеИзПланаРемонтов.СписокОбъектовРемонта КАК СписокОбъектовРемонта
	|ПОМЕСТИТЬ ДанныеИзОбъектовМаршрутов
	|ИЗ
	|	ДанныеИзПланаРемонтов КАК ДанныеИзПланаРемонтов
	|		ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.торо_МаршрутыРегламентныхМероприятий КАК торо_МаршрутыРегламентныхМероприятий
	|		ПО ДанныеИзПланаРемонтов.СписокОбъектовРемонта = торо_МаршрутыРегламентныхМероприятий.СписокОбъектов
	|
	|СГРУППИРОВАТЬ ПО
	|	торо_МаршрутыРегламентныхМероприятий.ОбъектРемонта,
	|	ДанныеИзПланаРемонтов.СписокОбъектовРемонта
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	ДанныеИзПланаРемонтов.ID КАК ID,
	|	ИСТИНА КАК Активность,
	|	ДанныеИзОбъектовМаршрутов.ОбъектРемонта КАК ОбъектРемонта,
	|	&Период КАК Период
	|ИЗ
	|	ДанныеИзПланаРемонтов КАК ДанныеИзПланаРемонтов
	|		ЛЕВОЕ СОЕДИНЕНИЕ ДанныеИзОбъектовМаршрутов КАК ДанныеИзОбъектовМаршрутов
	|		ПО ДанныеИзПланаРемонтов.СписокОбъектовРемонта = ДанныеИзОбъектовМаршрутов.СписокОбъектовРемонта";
	
	Запрос.УстановитьПараметр("ПланРемонтныхРабот", ПланРемонтовДвижения);
	Запрос.УстановитьПараметр("Период", Дата);
	
	ПланРемонтныхРабот = Запрос.Выполнить().Выгрузить();
		
	Движения.торо_ЗапланированныеМероприятияОбъектов.Записывать = Истина;
	Для Каждого СтрРемонта Из ПланРемонтныхРабот Цикл
		
		Движение = Движения.торо_ЗапланированныеМероприятияОбъектов.Добавить();
		ЗаполнитьЗначенияСвойств(Движение, СтрРемонта);
		
	КонецЦикла;
КонецПроцедуры

Процедура ДвиженияПоРегистру_торо_ПериодыПланированияППР(РежимПроведения, ДатаОкончанияПланирования, Отказ, ПланРемонтовДвижения)
	
	Если ВидОперации = Перечисления.торо_ВидыОперацийПланаГрафикаППР.ПланРабот Тогда
		
		НаборДвижений = Движения.торо_ПериодыПланированияППР;
		Движения.торо_ПериодыПланированияППР.Записывать = Истина;
		
		// Получим таблицу значений, совпадающую со структурой набора записей регистра.
		ТаблицаДвижений = НаборДвижений.Выгрузить();
		
		Запрос = Новый Запрос;
		Запрос.Текст = 
		"ВЫБРАТЬ
		|	ТабПланРемонтов.СписокОбъектовРемонта,
		|	ТабПланРемонтов.ВидРемонтныхРабот
		|ПОМЕСТИТЬ ТабПланРемонтов
		|ИЗ
		|	&ТабПланРемонтов КАК ТабПланРемонтов
		|;
		|
		|////////////////////////////////////////////////////////////////////////////////
		|ВЫБРАТЬ РАЗЛИЧНЫЕ
		|	ТабПланРемонтов.СписокОбъектовРемонта КАК ОбъектРемонтныхРабот,
		|	ТабПланРемонтов.ВидРемонтныхРабот
		|ИЗ
		|	ТабПланРемонтов КАК ТабПланРемонтов";
		
		Запрос.УстановитьПараметр("ТабПланРемонтов", ПланРемонтовДвижения);
		
		ТаблицаОбъектовРемонта = Запрос.Выполнить().Выгрузить();
		
		// Заполним таблицу движений.
		торо_ОбщегоНазначения.ЗагрузитьВТаблицуЗначений(ТаблицаОбъектовРемонта, ТаблицаДвижений);
		
		// Недостающие поля.
		ТаблицаДвижений.ЗаполнитьЗначения(ДатаПланирования, "ДатаПланирования");								  
		ТаблицаДвижений.ЗаполнитьЗначения(ПериодичностьДетализации, "ПериодичностьДетализации");
		ТаблицаДвижений.ЗаполнитьЗначения(ДатаОкончанияПланирования, "ДатаОкончанияПланирования");
		ТаблицаДвижений.ЗаполнитьЗначения(ЭтотОбъект.Ссылка, "Регистратор");
		ТаблицаДвижений.ЗаполнитьЗначения(ЭтотОбъект.Организация, "Организация");
		
		НаборДвижений.мТаблицаДвижений   = ТаблицаДвижений;
		
		Если Не Отказ Тогда
			Движения.торо_ПериодыПланированияППР.ВыполнитьДвижения();
		КонецЕсли;
		
	КонецЕсли;
	
КонецПроцедуры

Процедура ДвиженияПоРегистру_торо_ВсеДокументыПоРемонтам()
	Движения.торо_ВсеДокументыПоРемонтам.Записывать = Истина;
	Движения.торо_ВсеДокументыПоРемонтам.Очистить();
	
	Для каждого СтрокаТЧ Из ПланРемонтов Цикл
	    Движение = Движения.торо_ВсеДокументыПоРемонтам.Добавить();
		Движение.IDРемонта = СтрокаТЧ.ID;
		Движение.ДокументПоРемонту = Ссылка;
		Движение.ДатаДокумента = Дата;
		Движение.Организация = Организация;
		Движение.Подразделение = Подразделение;
	КонецЦикла;
КонецПроцедуры

Процедура ПроверитьПериодыПланированияРемонтныхРабот(Отказ)
	
	Запрос = Новый Запрос("ВЫБРАТЬ
	|	Маршруты.СписокОбъектовРемонта,
	|	Маршруты.ВидМероприятия
	|ПОМЕСТИТЬ Маршруты
	|ИЗ
	|	&Маршруты КАК Маршруты
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ РАЗЛИЧНЫЕ
	|	Маршруты.СписокОбъектовРемонта КАК ОбъектРемонтныхРабот,
	|	Маршруты.ВидМероприятия КАК ВидРемонтныхРабот,
	|	&ДатаПланирования КАК ДатаПланирования,
	|	&ПериодичностьДетализации КАК ПериодичностьДетализации,
	|	&ДатаОкончанияПланирования КАК ДатаОкончанияПланирования,
	|	&Регистратор КАК Регистратор,
	|	&Организация КАК Организация
	|ИЗ
	|	Маршруты КАК Маршруты");
	
	Запрос.УстановитьПараметр("Маршруты", Маршруты.Выгрузить());
	Запрос.УстановитьПараметр("ДатаПланирования", ДатаПланирования);								  
	Запрос.УстановитьПараметр("ПериодичностьДетализации", ПериодичностьДетализации);
	Если ЗначениеЗаполнено(ПериодичностьДетализации) Тогда		
		ДатаОкончанияПланирования = ОбщегоНазначенияУТКлиентСервер.РассчитатьДатуОкончанияПериода(ДатаПланирования, ПериодичностьДетализации, КоличествоПериодов);
	Иначе
		ТекстСообщения = НСтр("ru = 'Не выбран период планирования!'");
		ОбщегоНазначения.СообщитьПользователю(ТекстСообщения);
	КонецЕсли;
	Запрос.УстановитьПараметр("ДатаОкончанияПланирования", ДатаОкончанияПланирования);
	Запрос.УстановитьПараметр("Регистратор", ЭтотОбъект.Ссылка);
	Запрос.УстановитьПараметр("Организация", ЭтотОбъект.Организация);
	ТаблицаПроверкиПериодовПланирования = Запрос.Выполнить().Выгрузить();
	
	Запрос = Новый Запрос;
	Запрос.Текст = 
	"ВЫБРАТЬ
	|	ТаблицаПроверкиПериодовПланирования.ОбъектРемонтныхРабот КАК ОбъектРемонтныхРабот,
	|	ТаблицаПроверкиПериодовПланирования.ПериодичностьДетализации КАК ПериодичностьДетализации,
	|	ТаблицаПроверкиПериодовПланирования.ДатаПланирования КАК ДатаПланирования,
	|	ТаблицаПроверкиПериодовПланирования.ДатаОкончанияПланирования КАК ДатаОкончанияПланирования,
	|	ТаблицаПроверкиПериодовПланирования.ВидРемонтныхРабот КАК ВидРемонтныхРабот,
	|	ТаблицаПроверкиПериодовПланирования.Регистратор КАК Регистратор,
	|	ТаблицаПроверкиПериодовПланирования.Организация КАК Организация
	|ПОМЕСТИТЬ ТаблицаПроверкиПериодовПланирования
	|ИЗ
	|	&ТаблицаПроверкиПериодовПланирования КАК ТаблицаПроверкиПериодовПланирования
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	ПРЕДСТАВЛЕНИЕ(торо_ПериодыПланированияППР.ДатаПланирования) КАК ДатаПланированияПредставление,
	|	ПРЕДСТАВЛЕНИЕ(торо_ПериодыПланированияППР.ПериодичностьДетализации) КАК ПериодичностьДетализацииПредставление,
	|	ПРЕДСТАВЛЕНИЕ(торо_ПериодыПланированияППР.ОбъектРемонтныхРабот) КАК ОбъектРемонтныхРаботПредставление,
	|	ПРЕДСТАВЛЕНИЕ(торо_ПериодыПланированияППР.Регистратор) КАК РегистраторПредставление,
	|	ПРЕДСТАВЛЕНИЕ(торо_ПериодыПланированияППР.ВидРемонтныхРабот) КАК ВидРемонтныхРаботПредставление,
	|	торо_ПериодыПланированияППР.Регистратор КАК Регистратор,
	|	торо_ПериодыПланированияППР.ОбъектРемонтныхРабот КАК ОбъектРемонтныхРабот,
	|	торо_ПериодыПланированияППР.ВидРемонтныхРабот КАК ВидРемонтныхРабот,
	|	ТаблицаПроверкиПериодовПланирования.Организация КАК Организация,
	|	ТаблицаПроверкиПериодовПланирования.ДатаПланирования КАК ДатаПланирования,
	|	ТаблицаПроверкиПериодовПланирования.ДатаОкончанияПланирования КАК ДатаОкончанияПланирования
	|ПОМЕСТИТЬ ВТ_ПересекающиесяПериоды
	|ИЗ
	|	ТаблицаПроверкиПериодовПланирования КАК ТаблицаПроверкиПериодовПланирования
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ РегистрСведений.торо_ПериодыПланированияППР КАК торо_ПериодыПланированияППР
	|		ПО ТаблицаПроверкиПериодовПланирования.ОбъектРемонтныхРабот = торо_ПериодыПланированияППР.ОбъектРемонтныхРабот
	|			И ТаблицаПроверкиПериодовПланирования.ВидРемонтныхРабот = торо_ПериодыПланированияППР.ВидРемонтныхРабот
	|			И (ТаблицаПроверкиПериодовПланирования.ДатаОкончанияПланирования > торо_ПериодыПланированияППР.ДатаПланирования
	|					И ТаблицаПроверкиПериодовПланирования.ДатаОкончанияПланирования < торо_ПериодыПланированияППР.ДатаОкончанияПланирования
	|				ИЛИ ТаблицаПроверкиПериодовПланирования.ДатаПланирования > торо_ПериодыПланированияППР.ДатаПланирования
	|					И ТаблицаПроверкиПериодовПланирования.ДатаПланирования < торо_ПериодыПланированияППР.ДатаОкончанияПланирования
	|				ИЛИ ТаблицаПроверкиПериодовПланирования.ДатаПланирования < торо_ПериодыПланированияППР.ДатаПланирования
	|					И ТаблицаПроверкиПериодовПланирования.ДатаОкончанияПланирования > торо_ПериодыПланированияППР.ДатаПланирования
	|				ИЛИ ТаблицаПроверкиПериодовПланирования.ДатаОкончанияПланирования > торо_ПериодыПланированияППР.ДатаОкончанияПланирования
	|					И ТаблицаПроверкиПериодовПланирования.ДатаПланирования < торо_ПериодыПланированияППР.ДатаОкончанияПланирования)
	|			И ТаблицаПроверкиПериодовПланирования.Регистратор <> торо_ПериодыПланированияППР.Регистратор
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	торо_ПлановыеМероприятияСрезПоследних.СписокОбъектов КАК СписокОбъектов,
	|	торо_ПлановыеМероприятияСрезПоследних.ВидМероприятия КАК ВидМероприятия,
	|	торо_ПлановыеМероприятияСрезПоследних.ДатаНачала КАК ДатаНачала,
	|	торо_ПлановыеМероприятияСрезПоследних.ДатаОкончания КАК ДатаОкончания,
	|	торо_ПлановыеМероприятияСрезПоследних.Отменен КАК Отменен,
	|	торо_ПлановыеМероприятияСрезПоследних.Организация КАК Организация
	|ПОМЕСТИТЬ ВТ_ПлановыеРемонты
	|ИЗ
	|	РегистрСведений.торо_ПлановыеМероприятия.СрезПоследних(
	|			,
	|			(СписокОбъектов, ВидМероприятия) В
	|				(ВЫБРАТЬ
	|					ВТ_ПересекающиесяПериоды.ОбъектРемонтныхРабот КАК ОбъектРемонтныхРабот,
	|					ВТ_ПересекающиесяПериоды.ВидРемонтныхРабот КАК ВидРемонтныхРабот
	|				ИЗ
	|					ВТ_ПересекающиесяПериоды КАК ВТ_ПересекающиесяПериоды)) КАК торо_ПлановыеМероприятияСрезПоследних
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	ВТ_ПересекающиесяПериоды.ДатаПланированияПредставление КАК ДатаПланированияПредставление,
	|	ВТ_ПересекающиесяПериоды.ПериодичностьДетализацииПредставление КАК ПериодичностьДетализацииПредставление,
	|	ВТ_ПересекающиесяПериоды.ОбъектРемонтныхРаботПредставление КАК ОбъектРемонтныхРаботПредставление,
	|	ВТ_ПересекающиесяПериоды.РегистраторПредставление КАК РегистраторПредставление,
	|	ВТ_ПересекающиесяПериоды.ВидРемонтныхРаботПредставление КАК ВидРемонтныхРаботПредставление,
	|	ВТ_ПлановыеРемонты.Организация КАК Организация
	|ИЗ
	|	ВТ_ПересекающиесяПериоды КАК ВТ_ПересекающиесяПериоды
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ ВТ_ПлановыеРемонты КАК ВТ_ПлановыеРемонты
	|		ПО ВТ_ПересекающиесяПериоды.ОбъектРемонтныхРабот = ВТ_ПлановыеРемонты.СписокОбъектов
	|			И ВТ_ПересекающиесяПериоды.ВидРемонтныхРабот = ВТ_ПлановыеРемонты.ВидМероприятия
	|ГДЕ
	|	(ВТ_ПересекающиесяПериоды.Организация = ВТ_ПлановыеРемонты.Организация
	|			ИЛИ (ВТ_ПлановыеРемонты.Отменен = ЛОЖЬ
	|				ИЛИ ВТ_ПлановыеРемонты.Отменен ЕСТЬ NULL))
	|	И ВТ_ПлановыеРемонты.ДатаНачала МЕЖДУ ВТ_ПересекающиесяПериоды.ДатаПланирования И ВТ_ПересекающиесяПериоды.ДатаОкончанияПланирования
	|
	|СГРУППИРОВАТЬ ПО
	|	ВТ_ПересекающиесяПериоды.ДатаПланированияПредставление,
	|	ВТ_ПересекающиесяПериоды.ПериодичностьДетализацииПредставление,
	|	ВТ_ПересекающиесяПериоды.ОбъектРемонтныхРаботПредставление,
	|	ВТ_ПересекающиесяПериоды.РегистраторПредставление,
	|	ВТ_ПересекающиесяПериоды.ВидРемонтныхРаботПредставление,
	|	ВТ_ПлановыеРемонты.Организация";
	
	Запрос.УстановитьПараметр("ТаблицаПроверкиПериодовПланирования", ТаблицаПроверкиПериодовПланирования);
	
	Результат = Запрос.Выполнить();
	Если Не Результат.Пустой() Тогда
		
		Отказ = Истина;
		Выборка = Результат.Выбрать();
		Пока Выборка.Следующий() Цикл

			ШаблонСообщения = НСтр("ru = 'Для списка объектов ""%1"" и вида мероприятия ""%2"" на данный период составлен документ планирования: 
										|%3, с той же организацией (%4), или с не отменными ремонтами!'");
			
			ТекстСообщения = СтрШаблон(ШаблонСообщения, Выборка.ОбъектРемонтныхРаботПредставление, Выборка.ВидРемонтныхРаботПредставление, Выборка.РегистраторПредставление, Выборка.Организация);
			ОбщегоНазначения.СообщитьПользователю(ТекстСообщения);
			
		КонецЦикла;
		
	КонецЕсли; 
	
КонецПроцедуры

// Проверяет правильность заполнения строк табличной части "ПланРемонтов".
//
// Параметры:
//  Отказ                   - флаг отказа в проведении.
//  Заголовок               - строка, заголовок сообщения об ошибке проведения.
//
Процедура ПроверитьЗаполнениеТабличнойЧастиПланРемонтов(Отказ)
	
	// Укажем, что надо проверить:
	СтруктураОбязательныхПолей = Новый Структура("СписокОбъектовРемонта, ВидРемонтныхРабот, ДатаНач, ДатаКон, СпособВыполнения");
	
	// Вызовем общую процедуру проверки.
	торо_ЗаполнениеДокументов.ПроверитьЗаполнениеТабличнойЧасти(ЭтотОбъект, "ПланРемонтов", СтруктураОбязательныхПолей, Отказ, "");
	
	Если ЗначениеЗаполнено(ПериодичностьДетализации) Тогда		
		ДатаОкончанияПланирования = ОбщегоНазначенияУТКлиентСервер.РассчитатьДатуОкончанияПериода(ДатаПланирования, ПериодичностьДетализации, КоличествоПериодов);
	Иначе
		ТекстСообщения = НСтр("ru = 'Не выбран период планирования!'");
		ОбщегоНазначения.СообщитьПользователю(ТекстСообщения);
	КонецЕсли;
	
	
	ПредставлениеТабличнойЧасти = ЭтотОбъект.Метаданные().ТабличныеЧасти["ПланРемонтов"].Представление();
	ПредставлениеРеквизита      = ЭтотОбъект.Метаданные().ТабличныеЧасти["ПланРемонтов"].Реквизиты["Исполнитель"].Представление();
	ИспользоватьЛоты = ПолучитьФункциональнуюОпцию("торо_ИспользоватьЛоты");
	
	// Проверим принадлежность ремонтов периоду планирования и заполненность Исполнителей.
	Для Каждого СтрокаТЧ Из ПланРемонтов Цикл
		
		Если СтрокаТЧ.ДатаНач < ДатаПланирования Тогда
			
			ШаблонСообщения = НСтр("ru = 'В списке объектов ремонта: %1
				|Вид мероприятия: %2
				|дата начала работ %3 меньше даты начала планирования!'");
			ТекстСообщения = СтрШаблон(ШаблонСообщения, СтрокаТЧ.СписокОбъектовРемонта, СтрокаТЧ.ВидРемонтныхРабот, Формат(СтрокаТЧ.ДатаНач, "ДФ=dd.MM.yyyy"));
			Поле = СтрШаблон("Объект.ПланРемонтов[%1].ДатаНач", ПланРемонтов.Индекс(СтрокаТЧ)); 
			ОбщегоНазначения.СообщитьПользователю(ТекстСообщения,, Поле,, Отказ);
			
		ИначеЕсли СтрокаТЧ.ДатаНач > ДатаОкончанияПланирования Тогда
			
			ШаблонСообщения = НСтр("ru = 'В списке объектов ремонта: %1
				|Вид мероприятия: %2
				|дата начала работ %3 больше даты окончания планирования!'");
			ТекстСообщения = СтрШаблон(ШаблонСообщения, СтрокаТЧ.СписокОбъектовРемонта, СтрокаТЧ.ВидРемонтныхРабот, Формат(СтрокаТЧ.ДатаНач, "ДФ=dd.MM.yyyy"));
			Поле = СтрШаблон("Объект.ПланРемонтов[%1].ДатаНач", ПланРемонтов.Индекс(СтрокаТЧ)); 
			ОбщегоНазначения.СообщитьПользователю(ТекстСообщения,, Поле,, Отказ);
			
		ИначеЕсли СтрокаТЧ.ДатаНач > СтрокаТЧ.ДатаКон Тогда
			
			ШаблонСообщения = НСтр("ru = 'В списке объектов ремонта: %1
				|Вид мероприятия: %2
				|дата начала работ %3 больше даты окончания работ %4!'");
			ТекстСообщения = СтрШаблон(ШаблонСообщения, СтрокаТЧ.ОбъектРемонтныхРабот, СтрокаТЧ.ВидРемонтныхРабот, Формат(СтрокаТЧ.ДатаНач, "ДФ=dd.MM.yyyy"), Формат(СтрокаТЧ.ДатаКон, "ДФ=dd.MM.yyyy"));
			Поле = СтрШаблон("Объект.ПланРемонтов[%1].ДатаНач", ПланРемонтов.Индекс(СтрокаТЧ)); 
			ОбщегоНазначения.СообщитьПользователю(ТекстСообщения,, Поле,, Отказ);
		
		КонецЕсли; 
		
		Значение = СтрокаТЧ.Исполнитель;
		
		Если НЕ ЗначениеЗаполнено(Значение) И НЕ(СтрокаТЧ.СпособВыполнения = Перечисления.СпособыСтроительства.Подрядный И ИспользоватьЛоты) Тогда
			
			Если ПредставлениеТабличнойЧасти = Неопределено Тогда
				ПредставлениеТабличнойЧасти = ЭтотОбъект.Метаданные().ТабличныеЧасти["ПланРемонтов"].Представление();
			КонецЕсли;
			
			ШаблонСообщения = НСтр("ru = 'В строке номер ""%1"" табличной части ""%2"" не заполнено значение реквизита ""%3""!'");
			Поле = СтрШаблон("Объект.ПланРемонтов[%1].Исполнитель", ПланРемонтов.Индекс(СтрокаТЧ)); 
			ТекстСообщения = СтрШаблон(ШаблонСообщения, СокрЛП(СтрокаТЧ.НомерСтроки), ПредставлениеТабличнойЧасти, СокрЛП(ПредставлениеРеквизита));
			ОбщегоНазначения.СообщитьПользователю(ТекстСообщения,, Поле,, Отказ);
			
		КонецЕсли;
		
	КонецЦикла;
	
КонецПроцедуры

#КонецОбласти

#КонецЕсли