#Область ОбработчикиСобытийФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	
	ОткрыватьПриЗапуске = ОбщегоНазначения.ХранилищеОбщихНастроекЗагрузить("ОбщаяФорма.торо_ТекущиеЗадачиТОиР", "ОткрыватьПриЗапуске", Истина);
		
	Если Не ОткрыватьПриЗапуске Тогда
		Отказ = Истина;
		Возврат;
	КонецЕсли;
	
	Организация = ОбщегоНазначения.ХранилищеОбщихНастроекЗагрузить("НастройкиТОиР", "ОсновнаяОрганизация");
	ПодразделениеИсполнитель = ОбщегоНазначения.ХранилищеОбщихНастроекЗагрузить("НастройкиТОиР",	"ОсновноеПодразделение");
	
	ПоказыватьРемонтыППР = Истина;
	ПоказыватьРемонтыРегламентные = Истина;
	ПоказыватьДефекты = Истина;
	ПоказыватьВнешниеОснования = Истина;
	
	Если (торо_РаботаССогласованиями.ПроверитьИспользованиеСогласованияДокументов("", Ложь) 
		Или торо_РаботаССогласованиями.ПроверитьИспользованиеСогласованияДокументов("", Истина)
		Или торо_РаботаССогласованиями.ПроверитьИспользованиеСогласованияДокументов("", Ложь, Истина)) Тогда
		ОбновитьЗадачи();
	Иначе 
		Элементы.Согласование.Видимость = Ложь;
		Элементы.ТекущиеЗадачиТОИР.ОтображениеСтраниц = ОтображениеСтраницФормы.Нет;
	КонецЕсли;
		
	ОбновитьТекущиеРемонты();
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиКомандФормы

&НаКлиенте
Процедура Обновить(Команда)
	
	ОбновитьЗадачи();
	ОбновитьТекущиеРемонты();
	
КонецПроцедуры

&НаКлиенте
Процедура ОбновитьТаблицуРемонты(Команда)
	
	ОбновитьТекущиеРемонты();
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовШапкиФормы

&НаКлиенте
Процедура ОрганизацияПриИзменении(Элемент)
	ОбновитьЗадачи();
	ОбновитьТекущиеРемонты();
КонецПроцедуры

&НаКлиенте
Процедура ОткрыватьПриЗапускеПриИзменении(Элемент)
	
	ОбщегоНазначенияВызовСервера.ХранилищеОбщихНастроекСохранить("ОбщаяФорма.торо_ТекущиеЗадачиТОиР", "ОткрыватьПриЗапуске", ОткрыватьПриЗапуске);
	
КонецПроцедуры

&НаКлиенте
Процедура ПоказыватьРемонтыППРПриИзменении(Элемент)
	
	СформироватьЗаголовокСвернутойГруппы();
	
КонецПроцедуры

&НаКлиенте
Процедура ПоказыватьРемонтыРегламентныеПриИзменении(Элемент)
	
	СформироватьЗаголовокСвернутойГруппы();
	
КонецПроцедуры

&НаКлиенте
Процедура ПоказыватьДефектыПриИзменении(Элемент)
	
	СформироватьЗаголовокСвернутойГруппы();
	
КонецПроцедуры

&НаКлиенте
Процедура ПоказыватьВнешниеОснованияПриИзменении(Элемент)
	
	СформироватьЗаголовокСвернутойГруппы();
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовТаблицыФормыТекущиеЗадачиСписок

&НаКлиенте
Процедура ТекущиеЗадачиСписокВыбор(Элемент, ВыбраннаяСтрока, Поле, СтандартнаяОбработка)
	
	СтандартнаяОбработка = Ложь;
	
	ДанныеСтроки = Элементы.ТекущиеЗадачиСписок.ТекущиеДанные;
	
	Если Не ДанныеСтроки = Неопределено Тогда
		ПоказатьЗначение(,ДанныеСтроки.Ссылка);
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовТаблицыФормыТекущиеРемонты

&НаКлиенте
Процедура ТекущиеРемонтыВыбор(Элемент, ВыбраннаяСтрока, Поле, СтандартнаяОбработка)
	ТекущиеДанные = Элементы.ТекущиеРемонты.ТекущиеДанные;
	Если ТекущиеДанные = Неопределено Тогда
	    Возврат;
	КонецЕсли;
	
	ИмяРеквизита = СтрЗаменить(Поле.Имя, "ТекущиеРемонты", "");
	
	Если ИмяРеквизита = "ДатаНачала" ИЛИ ИмяРеквизита = "ДатаОкончания" ИЛИ ИмяРеквизита = "Статус" Тогда
	    ПоказатьЗначение(, ТекущиеДанные.ДокументОснование);
	Иначе
	    ПоказатьЗначение(, ТекущиеДанные[ИмяРеквизита]);
	 КонецЕсли;
	 
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

&НаСервере
Процедура ОбновитьЗадачи()
	
	ТаблицаСОтображаемымиДокументами = Новый ТаблицаЗначений;
	
	Запрос = Новый Запрос;
	Запрос.Текст =
	"ВЫБРАТЬ РАЗРЕШЕННЫЕ
	|	торо_СтатусыСогласованияДокументовРемонтныхРабот.СогласуемыйДокумент КАК Документ,
	|	торо_СтатусыСогласованияДокументовРемонтныхРабот.Порядок КАК Порядок,
	|	торо_СтатусыСогласованияДокументовРемонтныхРабот.Статус КАК Статус,
	|	торо_СтатусыСогласованияДокументовРемонтныхРабот.СогласуемыйДокумент.СпособСогласования КАК ДокументСпособСогласования,
	|	торо_СтатусыСогласованияДокументовРемонтныхРабот.НеНапоминать КАК НеНапоминать,
	|	торо_СтатусыСогласованияДокументовРемонтныхРабот.Пользователь КАК Пользователь
	|ПОМЕСТИТЬ ДокументыСоСпособамиСогласования
	|ИЗ
	|	РегистрСведений.торо_СтатусыСогласованияДокументовРемонтныхРабот КАК торо_СтатусыСогласованияДокументовРемонтныхРабот
	|ГДЕ
	|	(торо_СтатусыСогласованияДокументовРемонтныхРабот.Организация = &Организация
	|			ИЛИ &Организация = ЗНАЧЕНИЕ(Справочник.Организации.ПустаяСсылка))
	|	И торо_СтатусыСогласованияДокументовРемонтныхРабот.Дата = ДАТАВРЕМЯ(1, 1, 1, 0, 0, 0)
	|
	|ИНДЕКСИРОВАТЬ ПО
	|	ДокументСпособСогласования,
	|	Статус
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ РАЗРЕШЕННЫЕ РАЗЛИЧНЫЕ
	|	торо_МатрицаПереходаСтатусовДокументов.СпособСогласования КАК СпособСогласования,
	|	торо_МатрицаПереходаСтатусовДокументов.ТекущийСтатус КАК ТекущийСтатус,
	|	торо_МатрицаПереходаСтатусовДокументов.СледующийСтатус КАК СледующийСтатус,
	|	торо_НастройкиСтатусовСогласованияДокументов.НеОтображатьВСпискеЗадач КАК НеОтображатьВСпискеЗадач
	|ПОМЕСТИТЬ ВТ_ОтображаемыеСтатусы
	|ИЗ
	|	РегистрСведений.торо_МатрицаПереходаСтатусовДокументов КАК торо_МатрицаПереходаСтатусовДокументов
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ РегистрСведений.торо_НастройкиСтатусовСогласованияДокументов КАК торо_НастройкиСтатусовСогласованияДокументов
	|		ПО торо_МатрицаПереходаСтатусовДокументов.СпособСогласования = торо_НастройкиСтатусовСогласованияДокументов.СпособСогласования
	|			И торо_МатрицаПереходаСтатусовДокументов.СледующийСтатус = торо_НастройкиСтатусовСогласованияДокументов.Статус
	|ГДЕ
	|	(торо_МатрицаПереходаСтатусовДокументов.СпособСогласования, торо_МатрицаПереходаСтатусовДокументов.ТекущийСтатус) В
	|			(ВЫБРАТЬ
	|				ДокументыСоСпособамиСогласования.ДокументСпособСогласования,
	|				ДокументыСоСпособамиСогласования.Статус
	|			ИЗ
	|				ДокументыСоСпособамиСогласования КАК ДокументыСоСпособамиСогласования)
	|	И торо_НастройкиСтатусовСогласованияДокументов.НеОтображатьВСпискеЗадач = ЛОЖЬ
	|
	|ИНДЕКСИРОВАТЬ ПО
	|	СпособСогласования,
	|	ТекущийСтатус
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	ДокументыСоСпособамиСогласования.Документ КАК Документ,
	|	ДокументыСоСпособамиСогласования.Порядок КАК Порядок,
	|	ДокументыСоСпособамиСогласования.Статус КАК Статус,
	|	ДокументыСоСпособамиСогласования.ДокументСпособСогласования КАК ДокументСпособСогласования,
	|	ДокументыСоСпособамиСогласования.НеНапоминать КАК НеНапоминать,
	|	ДокументыСоСпособамиСогласования.Пользователь КАК Пользователь
	|ПОМЕСТИТЬ ВТ_ОтобранныеЗаписи
	|ИЗ
	|	ДокументыСоСпособамиСогласования КАК ДокументыСоСпособамиСогласования
	|ГДЕ
	|	(ДокументыСоСпособамиСогласования.ДокументСпособСогласования, ДокументыСоСпособамиСогласования.Статус) В
	|			(ВЫБРАТЬ
	|				ВТ_ОтображаемыеСтатусы.СпособСогласования КАК СпособСогласования,
	|				ВТ_ОтображаемыеСтатусы.ТекущийСтатус КАК ТекущийСтатус
	|			ИЗ
	|				ВТ_ОтображаемыеСтатусы КАК ВТ_ОтображаемыеСтатусы)
	|
	|ИНДЕКСИРОВАТЬ ПО
	|	Документ,
	|	Порядок
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ РАЗРЕШЕННЫЕ
	|	торо_СтатусыСогласованияДокументовРемонтныхРабот.Документ КАК Документ,
	|	МИНИМУМ(торо_СтатусыСогласованияДокументовРемонтныхРабот.Порядок) КАК Порядок
	|ПОМЕСТИТЬ ТекущийПорядок
	|ИЗ
	|	ВТ_ОтобранныеЗаписи КАК торо_СтатусыСогласованияДокументовРемонтныхРабот
	|
	|СГРУППИРОВАТЬ ПО
	|	торо_СтатусыСогласованияДокументовРемонтныхРабот.Документ
	|
	|ИНДЕКСИРОВАТЬ ПО
	|	Документ,
	|	Порядок
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ РАЗРЕШЕННЫЕ
	|	торо_СтатусыСогласованияДокументовРемонтныхРабот.Документ КАК Документ
	|ИЗ
	|	ТекущийПорядок КАК ТекущийПорядок
	|		ЛЕВОЕ СОЕДИНЕНИЕ ВТ_ОтобранныеЗаписи КАК торо_СтатусыСогласованияДокументовРемонтныхРабот
	|		ПО ТекущийПорядок.Документ = торо_СтатусыСогласованияДокументовРемонтныхРабот.Документ
	|			И ТекущийПорядок.Порядок = торо_СтатусыСогласованияДокументовРемонтныхРабот.Порядок
	|ГДЕ
	|	торо_СтатусыСогласованияДокументовРемонтныхРабот.Пользователь = &Пользователь
	|	И НЕ торо_СтатусыСогласованияДокументовРемонтныхРабот.НеНапоминать";
	
	Запрос.УстановитьПараметр("Организация", Организация);
	Запрос.УстановитьПараметр("Пользователь", Пользователи.ТекущийПользователь());
	
	Результат = Запрос.Выполнить();
	Выборка = Результат.Выбрать();
	
	ТекущиеЗадачиСписок.Очистить();
	
	МассивДобавляемых = ПроверитьВыборкуНаВозможностьДобавленияЗадач(Выборка);
	
	Для Каждого Документ Из МассивДобавляемых Цикл
		
		НоваяЗадача = ТекущиеЗадачиСписок.Добавить();
		НоваяЗадача.ТекстСписка	= "Согласовать документ " + Документ;
		НоваяЗадача.Ссылка = Документ;
		
	КонецЦикла;
	
	Элементы.ТекущиеЗадачиСписок.Обновить();
		
КонецПроцедуры

&НаСервере
Функция ПроверитьВыборкуНаВозможностьДобавленияЗадач(Выборка)
	
	МассивВозвращаемый = Новый Массив;
	
	Если Выборка.Количество() > 0 Тогда
		ФОИспользоватьМероприятия = ПолучитьФункциональнуюОпцию("торо_ИспользоватьРегламентныеМероприятия");
		ФОИспользоватьВнешниеОснования = ПолучитьФункциональнуюОпцию("торо_ИспользоватьВнешниеОснованияДляРабот");
		ФОИспользоватьДефекты = ПолучитьФункциональнуюОпцию("торо_УчетВыявленныхДефектовОборудования");
		ФОИспользоватьЗаявки = ПолучитьФункциональнуюОпцию("торо_ИспользоватьСметыРемонта");
		ФОИспользоватьНаряды = ПолучитьФункциональнуюОпцию("торо_ИспользоватьНарядыНаВыполнениеРабот");
		ФОИспользоватьПланГрафикРемонта = ПолучитьФункциональнуюОпцию("торо_ИспользоватьППР");
		ФОИспользоватьОстановочныеРемонты = ПолучитьФункциональнуюОпцию("торо_ИспользоватьОстановочныеРемонты");
		ФОИспользоватьАктПриемкиОборудования = ПолучитьФункциональнуюОпцию("торо_ИспользоватьАктПриемкиОборудования");
		ФОИспользоватьПланРаботПодразделения = ПолучитьФункциональнуюОпцию("торо_ИспользоватьОбщийПланРабот");
		ФОИспользоватьАнализКоренныхПричин = ПолучитьФункциональнуюОпцию("торо_ИспользоватьАнализКорневыхПричин");
		
		ВидДокумента = Перечисления.торо_ВидыДокументовСогласованияРемонтов.торо_АктОВыполненииЭтапаРабот;
		мИспользоватьСогласованияДляАктаОВыполненииРабот = торо_РаботаССогласованиями.ПроверитьИспользованиеСогласованияДокументов(ВидДокумента);
		
		ВидДокумента = Перечисления.торо_ВидыДокументовСогласованияРемонтов.торо_ВнешнееОснованиеДляРабот;
		мИспользоватьСогласованияДляВнешнихОснований = торо_РаботаССогласованиями.ПроверитьИспользованиеСогласованияДокументов(ВидДокумента);
		
		ВидДокумента = Перечисления.торо_ВидыДокументовСогласованияРемонтов.торо_ВыявленныеДефекты;
		мИспользоватьСогласованияДляДефектов = торо_РаботаССогласованиями.ПроверитьИспользованиеСогласованияДокументов(ВидДокумента);
		
		ВидДокумента = Перечисления.торо_ВидыДокументовСогласованияРемонтов.торо_ЗаявкаНаРемонт;
		мИспользоватьСогласованияДляЗаявок = торо_РаботаССогласованиями.ПроверитьИспользованиеСогласованияДокументов(ВидДокумента);
		
		ВидДокумента = Перечисления.торо_ВидыДокументовСогласованияРемонтов.торо_НарядНаВыполнениеРемонтныхРабот;
		мИспользоватьСогласованияДляНарядовНаРемРаботы = торо_РаботаССогласованиями.ПроверитьИспользованиеСогласованияДокументов(ВидДокумента);
		
		ВидДокумента = Перечисления.торо_ВидыДокументовСогласованияРемонтов.торо_ПланГрафикРемонта;
		мИспользоватьСогласованияДляПланГрафикРемонта = торо_РаботаССогласованиями.ПроверитьИспользованиеСогласованияДокументов(ВидДокумента);
		
		ВидДокумента = Перечисления.торо_ВидыДокументовСогласованияМероприятий.торо_АктОВыполненииРегламентногоМероприятия;
		мИспользоватьСогласованияДляАктовОВыполненииМероприятий = торо_РаботаССогласованиями.ПроверитьИспользованиеСогласованияДокументов(ВидДокумента, Истина);
		
		ВидДокумента = Перечисления.торо_ВидыДокументовСогласованияМероприятий.торо_ГрафикРегламентныхМероприятийТОиР;
		мИспользоватьСогласованияДляГрафиковМероприятий = торо_РаботаССогласованиями.ПроверитьИспользованиеСогласованияДокументов(ВидДокумента, Истина);
		
		ВидДокумента = Перечисления.торо_ВидыДокументовСогласованияМероприятий.торо_НарядНаРегламентноеМероприятие;
		мИспользоватьСогласованияДляНарядовНаМероприятия = торо_РаботаССогласованиями.ПроверитьИспользованиеСогласованияДокументов(ВидДокумента, Истина);
		
		ВидДокумента = Перечисления.торо_ВидыДокументовСогласованияРемонтов.торо_ОстановочныеРемонты;
		мИспользоватьСогласованияДляОстановочныхРемонтов = торо_РаботаССогласованиями.ПроверитьИспользованиеСогласованияДокументов(ВидДокумента);
		
		ВидДокумента = Перечисления.торо_ВидыДокументовСогласованияРемонтов.торо_АктПриемкиОборудования;
		мИспользоватьСогласованияДляАктовПриемкиОборудования = торо_РаботаССогласованиями.ПроверитьИспользованиеСогласованияДокументов(ВидДокумента);
		
		ВидДокумента = Перечисления.торо_ВидыДокументовСогласованияРемонтов.торо_ПланРаботПодразделения;
		мИспользоватьСогласованияДляПланРаботПодразделения = торо_РаботаССогласованиями.ПроверитьИспользованиеСогласованияДокументов(ВидДокумента);
		
		ВидДокумента = Перечисления.торо_ВидыДокументовСогласованияРемонтов.торо_АнализКоренныхПричин;
		мИспользоватьСогласованияДляАнализаКоренныхПричин = торо_РаботаССогласованиями.ПроверитьИспользованиеСогласованияДокументов(ВидДокумента);		
		
		ВидДокумента = Перечисления.торо_ВидыДокументовСогласованияМТО.торо_ЗаказПоставщику;
		ПолучатьЗаказПоставщику = торо_РаботаССогласованиями.ПроверитьИспользованиеСогласованияДокументов(ВидДокумента,, Истина);
		
		ВидДокумента = Перечисления.торо_ВидыДокументовСогласованияМТО.торо_УстановкаМинимальногоОстаткаНоменклатуры;
		ПолучатьУстановкуМинимальныхОстатков = торо_РаботаССогласованиями.ПроверитьИспользованиеСогласованияДокументов(ВидДокумента,, Истина);
		
		ПолучатьАктыОВыполненииЭтаповРабот = мИспользоватьСогласованияДляАктаОВыполненииРабот;
		ПолучатьАктыОВыполненииРегламентныхМероприятий = ФОИспользоватьМероприятия И мИспользоватьСогласованияДляАктовОВыполненииМероприятий;
		ПолучатьВнешниеОснованияДляРабот = ФОИспользоватьВнешниеОснования И мИспользоватьСогласованияДляВнешнихОснований;
		ПолучатьВыявленныеДефекты = ФОИспользоватьДефекты И мИспользоватьСогласованияДляДефектов;
		ПолучатьГрафикиРегламентныхМероприятий = ФОИспользоватьМероприятия И мИспользоватьСогласованияДляГрафиковМероприятий;
		ПолучатьЗаявкиНаРемонт = ФОИспользоватьЗаявки И мИспользоватьСогласованияДляЗаявок;
		ПолучатьНарядыНаВыполнениеРемонтныхРабот = ФОИспользоватьНаряды И мИспользоватьСогласованияДляНарядовНаРемРаботы;
		ПолучатьНарядыНаРегламентныеМероприятия = ФОИспользоватьМероприятия И мИспользоватьСогласованияДляНарядовНаМероприятия;
		ПолучатьПланыГрафикиРемонта = ФОИспользоватьПланГрафикРемонта и мИспользоватьСогласованияДляПланГрафикРемонта;
		ПолучатьОстановочныеРемонты = ФОИспользоватьОстановочныеРемонты и мИспользоватьСогласованияДляОстановочныхРемонтов;
		ПолучатьАктыПриемкиОборудования = ФОИспользоватьАктПриемкиОборудования и мИспользоватьСогласованияДляАктовПриемкиОборудования;
		ПолучатьПланРаботПодразделения = ФОИспользоватьПланРаботПодразделения и мИспользоватьСогласованияДляПланРаботПодразделения;
		ПолучатьАнализКоренныхПричин = ФОИспользоватьАнализКоренныхПричин и мИспользоватьСогласованияДляАнализаКоренныхПричин;
		
		Пока Выборка.Следующий() Цикл
			
			Документ = Выборка.Документ;
			
			// Проверка РЛС
			Если СтрНайти(Строка(Документ), НСтр("ru='Объект не найден'")) > 0 Тогда
				Продолжить;	
			КонецЕсли;
			
			Если ТипЗнч(Документ) = Тип("ДокументСсылка.торо_АктОВыполненииЭтапаРабот") 
					И ПолучатьАктыОВыполненииЭтаповРабот Тогда
				МассивВозвращаемый.Добавить(Документ);
			ИначеЕсли ТипЗнч(Документ) = Тип("ДокументСсылка.торо_АктОВыполненииРегламентногоМероприятия")
					И ПолучатьАктыОВыполненииРегламентныхМероприятий Тогда
				МассивВозвращаемый.Добавить(Документ);
			ИначеЕсли ТипЗнч(Документ) = Тип("ДокументСсылка.торо_ВнешнееОснованиеДляРабот")
					И ПолучатьВнешниеОснованияДляРабот Тогда
				МассивВозвращаемый.Добавить(Документ);
			ИначеЕсли ТипЗнч(Документ) = Тип("ДокументСсылка.торо_ВыявленныеДефекты") 
					И ПолучатьВыявленныеДефекты Тогда
				МассивВозвращаемый.Добавить(Документ);
			ИначеЕсли ТипЗнч(Документ) = Тип("ДокументСсылка.торо_ГрафикРегламентныхМероприятийТОиР") 
					И ПолучатьГрафикиРегламентныхМероприятий Тогда
				МассивВозвращаемый.Добавить(Документ);
			ИначеЕсли ТипЗнч(Документ) = Тип("ДокументСсылка.торо_ЗаявкаНаРемонт") 
					И ПолучатьЗаявкиНаРемонт Тогда
				МассивВозвращаемый.Добавить(Документ);
			ИначеЕсли ТипЗнч(Документ) = Тип("ДокументСсылка.торо_НарядНаВыполнениеРемонтныхРабот")
					И ПолучатьНарядыНаВыполнениеРемонтныхРабот Тогда
				МассивВозвращаемый.Добавить(Документ);
			ИначеЕсли ТипЗнч(Документ) = Тип("ДокументСсылка.торо_НарядНаРегламентноеМероприятие") 
					И ПолучатьНарядыНаРегламентныеМероприятия Тогда
				МассивВозвращаемый.Добавить(Документ);
			ИначеЕсли ТипЗнч(Документ) = Тип("ДокументСсылка.торо_ПланГрафикРемонта") 
					И ПолучатьПланыГрафикиРемонта Тогда
				МассивВозвращаемый.Добавить(Документ);
			ИначеЕсли ТипЗнч(Документ) = Тип("ДокументСсылка.торо_ОстановочныеРемонты") 
					И ПолучатьОстановочныеРемонты Тогда
				МассивВозвращаемый.Добавить(Документ);
			ИначеЕсли ТипЗнч(Документ) = Тип("ДокументСсылка.торо_АктПриемкиОборудования") 
					И ПолучатьАктыПриемкиОборудования Тогда
				МассивВозвращаемый.Добавить(Документ);
			ИначеЕсли ТипЗнч(Документ) = Тип("ДокументСсылка.торо_ПланРаботПодразделения") 
					И ПолучатьПланРаботПодразделения Тогда
				МассивВозвращаемый.Добавить(Документ);
			ИначеЕсли ТипЗнч(Документ) = Тип("ДокументСсылка.торо_ЗаказПоставщику") 
					И ПолучатьЗаказПоставщику Тогда
				МассивВозвращаемый.Добавить(Документ);
			ИначеЕсли ТипЗнч(Документ) = Тип("ДокументСсылка.торо_УстановкаМинимальногоОстаткаНоменклатуры") 
					И ПолучатьУстановкуМинимальныхОстатков Тогда
				МассивВозвращаемый.Добавить(Документ);
			ИначеЕсли ТипЗнч(Документ) = Тип("ДокументСсылка.торо_АнализКоренныхПричин") 
					И ПолучатьАнализКоренныхПричин Тогда
				МассивВозвращаемый.Добавить(Документ);				
			КонецЕсли;
			
		КонецЦикла;
	КонецЕсли;
	
	Возврат МассивВозвращаемый;
	
КонецФункции	

&НаСервере
Процедура ОбновитьТекущиеРемонты()
			
	Запрос = Новый Запрос();
	Запрос.Текст = "ВЫБРАТЬ РАЗРЕШЕННЫЕ
	               |	торо_АктуальныеПлановыеДатыРемонтов.IDРемонта КАК IDРемонта,
	               |	торо_АктуальныеПлановыеДатыРемонтов.ДатаНачала КАК ДатаНачала,
	               |	торо_АктуальныеПлановыеДатыРемонтов.ДатаОкончания КАК ДатаОкончания,
	               |	торо_АктуальныеПлановыеДатыРемонтов.КрайняяДатаЗавершения КАК КрайняяДатаЗавершения,
	               |	торо_АктуальныеПлановыеДатыРемонтов.ДокументОснование КАК ДокументОснование,
	               |	торо_АктуальныеПлановыеДатыРемонтов.ДокументНачалаЦепочки КАК ДокументНачалаЦепочки,
	               |	торо_АктуальныеПлановыеДатыРемонтов.ОбъектРемонта КАК ОбъектРемонта,
	               |	торо_ОбщиеДанныеПоРемонтам.ВидРемонта КАК ВидРемонта,
	               |	торо_ОбщиеДанныеПоРемонтам.Организация КАК Организация,
	               |	ЕСТЬNULL(торо_ПлановыеИсполнителиРемонтов.Исполнитель, НЕОПРЕДЕЛЕНО) КАК Исполнитель,
	               |	ЕСТЬNULL(торо_ПлановыеИсполнителиРемонтов.УточнениеИсполнителя, НЕОПРЕДЕЛЕНО) КАК УточнениеИсполнителя
	               |ПОМЕСТИТЬ ВТ_РемонтыСИсполнителями
	               |ИЗ
	               |	РегистрСведений.торо_АктуальныеПлановыеДатыРемонтов КАК торо_АктуальныеПлановыеДатыРемонтов
	               |		ВНУТРЕННЕЕ СОЕДИНЕНИЕ РегистрСведений.торо_ОбщиеДанныеПоРемонтам КАК торо_ОбщиеДанныеПоРемонтам
	               |		ПО торо_АктуальныеПлановыеДатыРемонтов.IDРемонта = торо_ОбщиеДанныеПоРемонтам.IDРемонта
	               |		ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.торо_ПлановыеИсполнителиРемонтов КАК торо_ПлановыеИсполнителиРемонтов
	               |		ПО торо_АктуальныеПлановыеДатыРемонтов.IDРемонта = торо_ПлановыеИсполнителиРемонтов.IDРемонта
	               |ГДЕ
	               |	(торо_ОбщиеДанныеПоРемонтам.ВидДокументаНачалаЦепочки = ЗНАЧЕНИЕ(Перечисление.торо_ВидыДокументовНачалаЦепочкиРемонтов.ВыявленныйДефект)
	               |				И &ПоказыватьДефекты
	               |			ИЛИ торо_ОбщиеДанныеПоРемонтам.ВидДокументаНачалаЦепочки = ЗНАЧЕНИЕ(Перечисление.торо_ВидыДокументовНачалаЦепочкиРемонтов.ВнешнееОснованиеДляРабот)
	               |				И &ПоказыватьВнешниеОснования
	               |			ИЛИ торо_ОбщиеДанныеПоРемонтам.ВидДокументаНачалаЦепочки = ЗНАЧЕНИЕ(Перечисление.торо_ВидыДокументовНачалаЦепочкиРемонтов.ГрафикРегламентныхМероприятий)
	               |				И &ПоказыватьРемонтыРегламентные
	               |			ИЛИ торо_ОбщиеДанныеПоРемонтам.ВидДокументаНачалаЦепочки = ЗНАЧЕНИЕ(Перечисление.торо_ВидыДокументовНачалаЦепочкиРемонтов.ПланГрафикППР)
	               |				И &ПоказыватьРемонтыППР)
	               |	И торо_АктуальныеПлановыеДатыРемонтов.ДатаНачала <= &ДатаНачала
	               |	И НЕ торо_ОбщиеДанныеПоРемонтам.Завершен
	               |	И НЕ торо_ОбщиеДанныеПоРемонтам.Отменен
	               |	И НЕ торо_ОбщиеДанныеПоРемонтам.Замещен
	               |	И торо_ОбщиеДанныеПоРемонтам.Организация = &Организация
	               |	И (&ПоВсемИсполнителям
	               |			ИЛИ торо_ПлановыеИсполнителиРемонтов.Исполнитель = &Исполнитель)
	               |
	               |ИНДЕКСИРОВАТЬ ПО
	               |	IDРемонта,
	               |	Исполнитель,
	               |	УточнениеИсполнителя
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	торо_СтатусыРемонтовВЗаявкахПоБригадамСрезПоследних.ЗаявкаНаРемонт КАК ЗаявкаНаРемонт,
	               |	торо_СтатусыРемонтовВЗаявкахПоБригадамСрезПоследних.ИДРемонта КАК ИДРемонта,
	               |	торо_СтатусыРемонтовВЗаявкахПоБригадамСрезПоследних.Исполнитель КАК Исполнитель,
	               |	торо_СтатусыРемонтовВЗаявкахПоБригадамСрезПоследних.УточнениеИсполнителя КАК УточнениеИсполнителя,
	               |	торо_СтатусыРемонтовВЗаявкахПоБригадамСрезПоследних.Статус КАК Статус
	               |ПОМЕСТИТЬ СтатусыРемонтовВЗаявкахПоБригадам
	               |ИЗ
	               |	РегистрСведений.торо_СтатусыРемонтовВЗаявкахПоБригадам.СрезПоследних(
	               |			,
	               |			ИДРемонта В
	               |				(ВЫБРАТЬ
	               |					ВТ_РемонтыСИсполнителями.IDРемонта КАК IDРемонта
	               |				ИЗ
	               |					ВТ_РемонтыСИсполнителями КАК ВТ_РемонтыСИсполнителями)) КАК торо_СтатусыРемонтовВЗаявкахПоБригадамСрезПоследних
	               |
	               |ИНДЕКСИРОВАТЬ ПО
	               |	ИДРемонта,
	               |	Исполнитель,
	               |	УточнениеИсполнителя
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	торо_СтатусыРемонтовСрезПоследних.IDРемонта КАК IDРемонта,
	               |	торо_СтатусыРемонтовСрезПоследних.Статус КАК Статус
	               |ПОМЕСТИТЬ СтатусыРемонтов
	               |ИЗ
	               |	РегистрСведений.торо_СтатусыРемонтов.СрезПоследних(
	               |			,
	               |			IDРемонта В
	               |				(ВЫБРАТЬ
	               |					ВТ_РемонтыСИсполнителями.IDРемонта КАК IDРемонта
	               |				ИЗ
	               |					ВТ_РемонтыСИсполнителями КАК ВТ_РемонтыСИсполнителями)) КАК торо_СтатусыРемонтовСрезПоследних
	               |
	               |ИНДЕКСИРОВАТЬ ПО
	               |	IDРемонта
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	ВТ_РемонтыСИсполнителями.IDРемонта КАК ID,
	               |	ВТ_РемонтыСИсполнителями.ДатаНачала КАК ДатаНачала,
	               |	ВТ_РемонтыСИсполнителями.ДатаОкончания КАК ДатаОкончания,
	               |	ВТ_РемонтыСИсполнителями.КрайняяДатаЗавершения КАК КрайняяДатаЗавершения,
	               |	ВТ_РемонтыСИсполнителями.ДокументОснование КАК ДокументОснование,
	               |	ВТ_РемонтыСИсполнителями.ДокументНачалаЦепочки КАК ДокументНачалаЦепочки,
	               |	ВТ_РемонтыСИсполнителями.ОбъектРемонта КАК ОбъектРемонта,
	               |	ВТ_РемонтыСИсполнителями.ВидРемонта КАК ВидРемонта,
	               |	ВТ_РемонтыСИсполнителями.Организация КАК Организация,
	               |	ВТ_РемонтыСИсполнителями.Исполнитель КАК Исполнитель,
	               |	ВТ_РемонтыСИсполнителями.УточнениеИсполнителя КАК УточнениеИсполнителя,
	               |	ЕСТЬNULL(СтатусыРемонтовВЗаявкахПоБригадам.Статус, СтатусыРемонтов.Статус) КАК Статус
	               |ИЗ
	               |	ВТ_РемонтыСИсполнителями КАК ВТ_РемонтыСИсполнителями
	               |		ЛЕВОЕ СОЕДИНЕНИЕ СтатусыРемонтовВЗаявкахПоБригадам КАК СтатусыРемонтовВЗаявкахПоБригадам
	               |		ПО ВТ_РемонтыСИсполнителями.IDРемонта = СтатусыРемонтовВЗаявкахПоБригадам.ИДРемонта
	               |			И ВТ_РемонтыСИсполнителями.Исполнитель = СтатусыРемонтовВЗаявкахПоБригадам.Исполнитель
	               |			И ВТ_РемонтыСИсполнителями.УточнениеИсполнителя = СтатусыРемонтовВЗаявкахПоБригадам.УточнениеИсполнителя
	               |		ЛЕВОЕ СОЕДИНЕНИЕ СтатусыРемонтов КАК СтатусыРемонтов
	               |		ПО ВТ_РемонтыСИсполнителями.IDРемонта = СтатусыРемонтов.IDРемонта
	               |
	               |УПОРЯДОЧИТЬ ПО
	               |	ДатаНачала,
	               |	ДатаОкончания";
						
	Запрос.УстановитьПараметр("ДатаНачала", КонецДня(ТекущаяДатаСеанса()));
	Запрос.УстановитьПараметр("Исполнитель", ПодразделениеИсполнитель);
	Запрос.УстановитьПараметр("ПоВсемИсполнителям", НЕ ЗначениеЗаполнено(ПодразделениеИсполнитель));
	Запрос.УстановитьПараметр("Организация", Организация);
	Запрос.УстановитьПараметр("ПоказыватьВнешниеОснования", ПоказыватьВнешниеОснования);
	Запрос.УстановитьПараметр("ПоказыватьДефекты", ПоказыватьДефекты);
	Запрос.УстановитьПараметр("ПоказыватьРемонтыППР", ПоказыватьРемонтыППР);
	Запрос.УстановитьПараметр("ПоказыватьРемонтыРегламентные", ПоказыватьРемонтыРегламентные);
	
	РезультатЗапроса = Запрос.Выполнить();
	Выборка = РезультатЗапроса.Выбрать();
	
	ТекущиеРемонты.Очистить();
	
	Пока Выборка.Следующий() Цикл
		НоваяСтрока = ТекущиеРемонты.Добавить();
		ЗаполнитьЗначенияСвойств(НоваяСтрока, Выборка);
	КонецЦикла;
	
КонецПроцедуры

&НаКлиенте
Процедура СформироватьЗаголовокСвернутойГруппы()
	
	ЗаголовокСвернутойГруппы = НСтр("ru = 'Выводить информацию по%1'");
	
	МассивСтрокДокументов = Новый Массив;
	
	Если ПоказыватьРемонтыППР Тогда
		МассивСтрокДокументов.Добавить(НСтр("ru = ' план-графикам ППР'"));
	КонецЕсли;
	
	Если ПоказыватьРемонтыРегламентные Тогда
		МассивСтрокДокументов.Добавить(НСтр("ru = ' графикам регламентных мероприятий'"));
	КонецЕсли;
	
	Если ПоказыватьДефекты Тогда
		МассивСтрокДокументов.Добавить(НСтр("ru = ' дефектам'"));
	КонецЕсли;
	
	Если ПоказыватьВнешниеОснования Тогда
		МассивСтрокДокументов.Добавить(НСтр("ru = ' внешним основаниям'"));
	КонецЕсли;

	Элементы.ГруппаВыводитьИнформациюПо.ЗаголовокСвернутогоОтображения = СтрШаблон(ЗаголовокСвернутойГруппы, СтрСоединить(МассивСтрокДокументов, ","));
	
КонецПроцедуры

#КонецОбласти
