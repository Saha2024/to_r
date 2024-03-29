
//++ Проф-ИТ, #128, Башинская А., 25.08.2023

#Область ОбработчикиСобытийФормы

&НаКлиенте
Процедура ПриОткрытии(Отказ) 
   
   ПериодОтчета.Вариант = ВариантСтандартногоПериода.Месяц;
   ПериодОтчета.ДатаНачала = НачалоМесяца(ОбщегоНазначенияКлиент.ДатаСеанса());     
   ПериодОтчета.ДатаОкончания = КонецМесяца(ОбщегоНазначенияКлиент.ДатаСеанса()); 
   
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиКомандФормы

&НаКлиенте
Процедура Сформировать(Команда)
   СформироватьНаСервере();
КонецПроцедуры 

&НаКлиенте
Процедура Печать(Команда)
   
   Идентификатор = "ИД_ПечатьЖурналУчетаЗаявочныхРемонтовЗаМесяц";
	   
   КоллекцияПечатныхФорм = УправлениеПечатьюКлиент.НоваяКоллекцияПечатныхФорм(Идентификатор);
   ОписаниеПечатнойФормы = УправлениеПечатьюКлиент.ОписаниеПечатнойФормы(КоллекцияПечатныхФорм, Идентификатор);
   ОписаниеПечатнойФормы.ТабличныйДокумент = ПолучитьДокументДляПечати();
   
   УправлениеПечатьюКлиент.ПечатьДокументов(КоллекцияПечатныхФорм, , ЭтаФорма);
   
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

&НаСервере
Процедура СформироватьНаСервере()
   ВывестиОтчет();
КонецПроцедуры

&НаСервере
Процедура ВывестиОтчет()
   
   Табл = ТЧОтчет; 
   Табл.Очистить();
   ЭтаОбработка = РеквизитФормыВЗначение("Отчет");
   Макет = ЭтаОбработка.ПолучитьМакет("Макет");
   ОбластьШапка = Макет.ПолучитьОбласть("Шапка");
   ОбластьШапка2 = Макет.ПолучитьОбласть("Шапка2");
   
   Запрос = Новый Запрос;
   Запрос.УстановитьПараметр("ДатаНачала", НачалоМесяца(ПериодОтчета.ДатаНачала)); 
   Запрос.УстановитьПараметр("ДатаОкончания", КонецМесяца(ПериодОтчета.ДатаОкончания));   
   Запрос.Текст = ТекстЗапросаПечатьЖурналУчетаЗаявочныхРемонтовЗаМесяц(ОбъектРемонта, Запрос);        
   РезультатЗапроса = Запрос.ВыполнитьПакет();  
   
   ВыборкаОбъектыРемонта = РезультатЗапроса[5].Выбрать();
   ВыборкаНаработка = РезультатЗапроса[3].Выбрать();
   ВыборкаНоменклатура = РезультатЗапроса[4].Выбрать();
   
   Если ВыборкаОбъектыРемонта.Количество() > 0 Тогда	
	   Табл.Вывести(ОбластьШапка); 
   КонецЕсли;
   
   Пока ВыборкаОбъектыРемонта.Следующий() Цикл
	   ОбластьСтрока = Макет.ПолучитьОбласть("ОбластьСтрока");
	   ОбластьСтрока.Параметры.Заполнить(ВыборкаОбъектыРемонта); 
	   ОбластьСтрока.Параметры.Наименование = ВыборкаОбъектыРемонта.Наименование + " " + ВыборкаОбъектыРемонта.Марка;
	   ОбластьСтрока.Параметры.Пробег = "";
	   ОбластьСтрока.Параметры.Наработка = "";
	   
	   Отбор = Новый Структура;
	   Отбор.Вставить("ОбъектРемонта", ВыборкаОбъектыРемонта.ОбъектРемонта);
	   Отбор.Вставить("ID", ВыборкаОбъектыРемонта.ID);
	   
	   ЗаполнитьПробег(Отбор, ВыборкаНаработка, ОбластьСтрока);
	   
	   ОбластьСтрока.Параметры.ВремяРаботы = ?(ЗначениеЗаполнено(ВыборкаОбъектыРемонта.ВремяРаботы),
												   Строка(ВыборкаОбъектыРемонта.ВремяРаботы) + ", ч",
												   "");  
	   ОбластьСтрока.Параметры.ВремяПростоя = ?(ЗначениеЗаполнено(ВыборкаОбъектыРемонта.ВремяПростоя),
												   Строка(ВыборкаОбъектыРемонта.ВремяПростоя) + ", ч",
												   "");
	   Табл.Вывести(ОбластьСтрока); 
	   
	   ВывестиВыборкуНоменклатуры(ВыборкаОбъектыРемонта, ВыборкаНоменклатура, ОбластьШапка2, Макет, Табл);
			   
   КонецЦикла;
		   
КонецПроцедуры 

&НаСервере
Процедура ВывестиВыборкуНоменклатуры(ВыборкаОбъектыРемонта, ВыборкаНоменклатура, ОбластьШапка2, Макет, ТабДок)
   
   Отбор = Новый Структура;
   Отбор.Вставить("ОбъектРемонта", ВыборкаОбъектыРемонта.ОбъектРемонта);
   Отбор.Вставить("ID", ВыборкаОбъектыРемонта.ID);
   ШапкаВыведена = Ложь;
   Пока ВыборкаНоменклатура.НайтиСледующий(Отбор) Цикл
	   
	   Если ШапкаВыведена = Ложь Тогда
		   ТабДок.Вывести(ОбластьШапка2);
		   ШапкаВыведена = Истина;
	   КонецЕсли;
	   
	   ОбластьСтрока2 = Макет.ПолучитьОбласть("ОбластьСтрока2");
	   ОбластьСтрока2.Параметры.Заполнить(ВыборкаНоменклатура);
	   ТабДок.Вывести(ОбластьСтрока2);	
   КонецЦикла;	
   ВыборкаНоменклатура.Сбросить();
   
КонецПроцедуры

&НаСервере
Функция ТекстЗапросаПечатьЖурналУчетаЗаявочныхРемонтовЗаМесяц(ОбъектРемонта, Запрос)
   
   ТекстЗапроса = "ВЫБРАТЬ
                  |	торо_ВыявленныеДефектыСписокДефектов.ОбъектРемонта КАК ОбъектРемонта,
                  |	торо_ВыявленныеДефектыСписокДефектов.ID КАК ID,
                  |	ЕСТЬNULL(торо_ОбъектыРемонта.ИнвентарныйНомер, """") КАК ИнвентарныйНомер,
                  |	ЕСТЬNULL(торо_ОбъектыРемонта.Наименование, """") КАК Наименование,
                  |	ЕСТЬNULL(торо_ОбъектыРемонта.проф_Марка, """") КАК проф_Марка,
                  |	ЕСТЬNULL(торо_ОбъектыРемонта.ЗаводскойНомер, """") КАК ЗаводскойНомер,
//++ Проф-ИТ, #296, Горетовская М.С., 05.10.2023 - Доработки по задаче #128
                  |	ЕСТЬNULL(торо_ОбъектыРемонта.проф_Госномер, """") КАК ГосНомер,
//-- Проф-ИТ, #296, Горетовская М.С., 05.10.2023 - Доработки по задаче #128
                  |	ЕСТЬNULL(торо_ЗаявкаНаРемонтРемонтыОборудования.ДатаНачала, ДАТАВРЕМЯ(1, 1, 1)) КАК ДатаНачала,
                  |	ЕСТЬNULL(торо_ЗаявкаНаРемонтРемонтыОборудования.ДатаОкончания, ДАТАВРЕМЯ(1, 1, 1)) КАК ДатаОкончания,
                  |	торо_ВыявленныеДефектыСписокДефектов.ДефектОписание КАК ДефектОписание,
                  |	ЕСТЬNULL(торо_ОбъектыРемонта.ПлановыйГрафикРаботы, """") КАК ПлановыйГрафикРаботы,
                  |	торо_ВыявленныеДефектыСписокДефектов.ДефектПричина КАК ДефектПричина,
                  |	ЕСТЬNULL(торо_ЗаявкаНаРемонтРемонтыОборудования.Ссылка, ЗНАЧЕНИЕ(Документ.торо_ЗаявкаНаРемонт.ПустаяСсылка)) КАК ЗаявкаНаРемонт,
                  |	СУММА(ВЫРАЗИТЬ(торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.ВремяРаботы / 3600 КАК ЧИСЛО(12, 2))) КАК ВремяРаботы,
                  |	ЕСТЬNULL(торо_Наработка.Значение, 0) КАК Наработка,
                  |	ЕСТЬNULL(торо_Пробег.Значение, 0) КАК Пробег
                  |ПОМЕСТИТЬ ВТ_Предв
                  |ИЗ
                  |	Документ.торо_ВыявленныеДефекты.СписокДефектов КАК торо_ВыявленныеДефектыСписокДефектов
                  |		ВНУТРЕННЕЕ СОЕДИНЕНИЕ Документ.торо_ВыявленныеДефекты КАК торо_ВыявленныеДефекты
                  |			ЛЕВОЕ СОЕДИНЕНИЕ Документ.торо_ВыявленныеДефекты.ДополнительныеРеквизиты КАК торо_Пробег
                  |			ПО торо_ВыявленныеДефекты.Ссылка = торо_Пробег.Ссылка
                  |				И (торо_Пробег.Свойство.Наименование = ""Пробег"")
                  |			ЛЕВОЕ СОЕДИНЕНИЕ Документ.торо_ВыявленныеДефекты.ДополнительныеРеквизиты КАК торо_Наработка
                  |			ПО торо_ВыявленныеДефекты.Ссылка = торо_Наработка.Ссылка
                  |				И (торо_Наработка.Свойство.Наименование = ""Наработка"")
                  |		ПО торо_ВыявленныеДефектыСписокДефектов.Ссылка = торо_ВыявленныеДефекты.Ссылка
                  |			И (торо_ВыявленныеДефекты.ПометкаУдаления = ЛОЖЬ)
                  |			И (торо_ВыявленныеДефекты.Проведен)
                  |			И (&ОбъектРемонта)
                  |		ЛЕВОЕ СОЕДИНЕНИЕ Справочник.торо_ОбъектыРемонта КАК торо_ОбъектыРемонта
                  |		ПО торо_ВыявленныеДефектыСписокДефектов.ОбъектРемонта = торо_ОбъектыРемонта.Ссылка
                  |		ЛЕВОЕ СОЕДИНЕНИЕ Документ.торо_ЗаявкаНаРемонт.РемонтыОборудования КАК торо_ЗаявкаНаРемонтРемонтыОборудования
                  |		ПО (торо_ЗаявкаНаРемонтРемонтыОборудования.ID = торо_ВыявленныеДефектыСписокДефектов.ID)
                  |			И (торо_ОбъектыРемонта.Ссылка = торо_ЗаявкаНаРемонтРемонтыОборудования.ОбъектРемонта)
                  |			И (торо_ЗаявкаНаРемонтРемонтыОборудования.Ссылка.Проведен)
                  |		ЛЕВОЕ СОЕДИНЕНИЕ Документ.торо_АктОВыполненииЭтапаРабот.ТрудовыеЗатраты КАК торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты
                  |		ПО торо_ВыявленныеДефектыСписокДефектов.ID = торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.РемонтыОборудования_ID
                  |			И (торо_АктОВыполненииЭтапаРаботТрудовыеЗатраты.Ссылка.Проведен)
                  |ГДЕ
                  |	НАЧАЛОПЕРИОДА(торо_ЗаявкаНаРемонтРемонтыОборудования.ДатаНачала, МЕСЯЦ) = &ДатаНачала
                  |
                  |СГРУППИРОВАТЬ ПО
                  |	торо_ВыявленныеДефектыСписокДефектов.ОбъектРемонта,
                  |	торо_ВыявленныеДефектыСписокДефектов.ID,
                  |	ЕСТЬNULL(торо_ОбъектыРемонта.ИнвентарныйНомер, """"),
                  |	ЕСТЬNULL(торо_ОбъектыРемонта.Наименование, """"),
                  |	ЕСТЬNULL(торо_ОбъектыРемонта.проф_Марка, """"),
                  |	ЕСТЬNULL(торо_ОбъектыРемонта.ЗаводскойНомер, """"),
 //++ Проф-ИТ, #296, Горетовская М.С., 05.10.2023 - Доработки по задаче #128
                  |	ЕСТЬNULL(торо_ОбъектыРемонта.проф_Госномер, """"),
//-- Проф-ИТ, #296, Горетовская М.С., 05.10.2023 - Доработки по задаче #128
                 |	торо_ЗаявкаНаРемонтРемонтыОборудования.Ссылка,
                  |	ЕСТЬNULL(торо_ЗаявкаНаРемонтРемонтыОборудования.ДатаНачала, ДАТАВРЕМЯ(1, 1, 1)),
                  |	ЕСТЬNULL(торо_ЗаявкаНаРемонтРемонтыОборудования.ДатаОкончания, ДАТАВРЕМЯ(1, 1, 1)),
                  |	торо_ВыявленныеДефектыСписокДефектов.ДефектОписание,
                  |	торо_ОбъектыРемонта.ПлановыйГрафикРаботы,
                  |	торо_ВыявленныеДефектыСписокДефектов.ДефектПричина,
                  |	ЕСТЬNULL(торо_Наработка.Значение, 0),
                  |	ЕСТЬNULL(торо_Пробег.Значение, 0)
                  |
                  |ИНДЕКСИРОВАТЬ ПО
                  |	ОбъектРемонта,
                  |	ДатаНачала
                  |;
                  |
                  |////////////////////////////////////////////////////////////////////////////////
                  |ВЫБРАТЬ
                  |	ВТ_Предв.ОбъектРемонта КАК ОбъектРемонта,
                  |	ВТ_Предв.ID КАК ID,
                  |	ВТ_Предв.ИнвентарныйНомер КАК ИнвентарныйНомер,
                  |	ВТ_Предв.Наименование КАК Наименование,
                  |	ВТ_Предв.проф_Марка КАК проф_Марка,
                  |	ВТ_Предв.ЗаводскойНомер КАК ЗаводскойНомер,
//++ Проф-ИТ, #296, Горетовская М.С., 05.10.2023 - Доработки по задаче #128
                  |	ВТ_Предв.ГосНомер КАК ГосНомер,
//-- Проф-ИТ, #296, Горетовская М.С., 05.10.2023 - Доработки по задаче #128
                  |	ВТ_Предв.ЗаявкаНаРемонт КАК ЗаявкаНаРемонт,
                  |	ВТ_Предв.ДатаНачала КАК ДатаНачала,
                  |	ВТ_Предв.ДатаОкончания КАК ДатаОкончания,
                  |	ВТ_Предв.ДефектОписание КАК ДефектОписание,
                  |	ВТ_Предв.ПлановыйГрафикРаботы КАК ПлановыйГрафикРаботы,
                  |	ВТ_Предв.ДефектПричина КАК ДефектПричина,
                  |	ВТ_Предв.ВремяРаботы КАК ВремяРаботы,
                  |	СУММА(РАЗНОСТЬДАТ(ЕСТЬNULL(торо_АктОВыполненииЭтапаРаботРемонтыОборудования.ДатаНачала, ДАТАВРЕМЯ(1, 1, 1)), ЕСТЬNULL(торо_АктОВыполненииЭтапаРаботРемонтыОборудования.ДатаОкончания, ДАТАВРЕМЯ(1, 1, 1)), ЧАС)) КАК ВремяПростоя,
                  |	ВТ_Предв.Наработка КАК Наработка,
                  |	ВТ_Предв.Пробег КАК Пробег
                  |ПОМЕСТИТЬ ВТ_Итог
                  |ИЗ
                  |	ВТ_Предв КАК ВТ_Предв
                  |		ЛЕВОЕ СОЕДИНЕНИЕ Документ.торо_АктОВыполненииЭтапаРабот.РемонтыОборудования КАК торо_АктОВыполненииЭтапаРаботРемонтыОборудования
                  |		ПО (торо_АктОВыполненииЭтапаРаботРемонтыОборудования.ID = ВТ_Предв.ID)
                  |			И ВТ_Предв.ОбъектРемонта = торо_АктОВыполненииЭтапаРаботРемонтыОборудования.ОбъектРемонта
                  |
                  |СГРУППИРОВАТЬ ПО
                  |	ВТ_Предв.ОбъектРемонта,
                  |	ВТ_Предв.ID,
                  |	ВТ_Предв.ИнвентарныйНомер,
                  |	ВТ_Предв.Наименование,
                  |	ВТ_Предв.ЗаявкаНаРемонт,
                  |	ВТ_Предв.проф_Марка,
                  |	ВТ_Предв.ЗаводскойНомер,
//++ Проф-ИТ, #296, Горетовская М.С., 05.10.2023 - Доработки по задаче #128
                  |	ВТ_Предв.ГосНомер,
//-- Проф-ИТ, #296, Горетовская М.С., 05.10.2023 - Доработки по задаче #128
                  |	ВТ_Предв.ДатаНачала,
                  |	ВТ_Предв.ДатаОкончания,
                  |	ВТ_Предв.ДефектОписание,
                  |	ВТ_Предв.ПлановыйГрафикРаботы,
                  |	ВТ_Предв.ДефектПричина,
                  |	ВТ_Предв.ВремяРаботы,
                  |	ВТ_Предв.Наработка,
                  |	ВТ_Предв.Пробег
                  |
                  |ИНДЕКСИРОВАТЬ ПО
                  |	ОбъектРемонта,
                  |	ID
                  |;
                  |
                  |////////////////////////////////////////////////////////////////////////////////
                  |ВЫБРАТЬ
                  |	ВТ_Итог.ОбъектРемонта КАК ОбъектРемонта,
                  |	ВТ_Итог.ID КАК ID,
                  |	ВТ_Итог.ИнвентарныйНомер КАК ИнвентарныйНомер,
                  |	ВТ_Итог.Наименование КАК Наименование,
                  |	ВТ_Итог.проф_Марка КАК проф_Марка,
                  |	ВТ_Итог.ЗаводскойНомер КАК ЗаводскойНомер,
                  |	ВТ_Итог.ЗаявкаНаРемонт КАК ЗаявкаНаРемонт,
                  |	ВТ_Итог.ДатаНачала КАК ДатаНачала,
                  |	ВТ_Итог.ДатаОкончания КАК ДатаОкончания,
                  |	ВТ_Итог.ДефектОписание КАК ДефектОписание,
                  |	ВТ_Итог.ПлановыйГрафикРаботы КАК ПлановыйГрафикРаботы,
                  |	ВТ_Итог.ДефектПричина КАК ДефектПричина,
                  |	ВТ_Итог.ВремяРаботы КАК ВремяРаботы,
                  |	ВТ_Итог.ВремяПростоя КАК ВремяПростоя,
                  |	ВТ_Итог.Наработка КАК Наработка,
                  |	ВТ_Итог.Пробег КАК Пробег
                  |ПОМЕСТИТЬ ВТ_ИтогСНаработками
                  |ИЗ
                  |	ВТ_Итог КАК ВТ_Итог
                  |
                  |СГРУППИРОВАТЬ ПО
                  |	ВТ_Итог.ОбъектРемонта,
                  |	ВТ_Итог.ID,
                  |	ВТ_Итог.ИнвентарныйНомер,
                  |	ВТ_Итог.Наименование,
                  |	ВТ_Итог.ЗаявкаНаРемонт,
                  |	ВТ_Итог.проф_Марка,
                  |	ВТ_Итог.ЗаводскойНомер,
                  |	ВТ_Итог.ДатаНачала,
                  |	ВТ_Итог.ДатаОкончания,
                  |	ВТ_Итог.ДефектОписание,
                  |	ВТ_Итог.ПлановыйГрафикРаботы,
                  |	ВТ_Итог.ДефектПричина,
                  |	ВТ_Итог.ВремяРаботы,
                  |	ВТ_Итог.ВремяПростоя,
                  |	ВТ_Итог.Наработка,
                  |	ВТ_Итог.Пробег
                  |
                  |ИНДЕКСИРОВАТЬ ПО
                  |	ОбъектРемонта,
                  |	ID
                  |;
                  |
                  |////////////////////////////////////////////////////////////////////////////////
                  |ВЫБРАТЬ
                  |	ВТ_ИтогСНаработками.ОбъектРемонта КАК ОбъектРемонта,
                  |	ВТ_ИтогСНаработками.ID КАК ID,
                  |	ЕСТЬNULL(ВТ_ИтогСНаработками.Наработка, 0) КАК Наработка,
                  |	ЕСТЬNULL(ВТ_ИтогСНаработками.Пробег, 0) КАК Пробег
                  |ИЗ
                  |	ВТ_ИтогСНаработками КАК ВТ_ИтогСНаработками
                  |
                  |СГРУППИРОВАТЬ ПО
                  |	ВТ_ИтогСНаработками.ОбъектРемонта,
                  |	ВТ_ИтогСНаработками.ID,
                  |	ВТ_ИтогСНаработками.Наработка,
                  |	ВТ_ИтогСНаработками.Пробег
                  |
                  |УПОРЯДОЧИТЬ ПО
                  |	ОбъектРемонта
                  |;
                  |
                  |////////////////////////////////////////////////////////////////////////////////
                  |ВЫБРАТЬ
                  |	ЕСТЬNULL(ВнутреннееПотреблениеТоваровТовары.Номенклатура, ЗНАЧЕНИЕ(Справочник.Номенклатура.ПустаяСсылка)) КАК Номенклатура,
                  |	СУММА(ЕСТЬNULL(ВнутреннееПотреблениеТоваровТовары.Количество, 0)) КАК Количество,
                  |	ВТ_Итог.ОбъектРемонта КАК ОбъектРемонта,
                  |	ВТ_Итог.ID КАК ID
                  |ИЗ
                  |	ВТ_Итог КАК ВТ_Итог
                  |		ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.торо_ИнтеграцияДокументов КАК торо_ИнтеграцияДокументов
                  |		ПО ВТ_Итог.ID = торо_ИнтеграцияДокументов.ID
                  |			И (ТИПЗНАЧЕНИЯ(торо_ИнтеграцияДокументов.ДокументЕРП) = ТИП(Документ.ВнутреннееПотреблениеТоваров))
                  |			И (НЕ торо_ИнтеграцияДокументов.ДокументЕРП.ПометкаУдаления)
                  |		ЛЕВОЕ СОЕДИНЕНИЕ Документ.ВнутреннееПотреблениеТоваров.Товары КАК ВнутреннееПотреблениеТоваровТовары
                  |		ПО (торо_ИнтеграцияДокументов.ДокументЕРП = ВнутреннееПотреблениеТоваровТовары.Ссылка)
                  |			И (ВнутреннееПотреблениеТоваровТовары.Ссылка.Проведен)
                  |ГДЕ
                  |	ВнутреннееПотреблениеТоваровТовары.Количество <> 0
                  |
                  |СГРУППИРОВАТЬ ПО
                  |	ВнутреннееПотреблениеТоваровТовары.Номенклатура,
                  |	ВТ_Итог.ОбъектРемонта,
                  |	ВТ_Итог.ID
                  |;
                  |
                  |////////////////////////////////////////////////////////////////////////////////
                  |ВЫБРАТЬ
                  |	ВТ_Итог.ОбъектРемонта КАК ОбъектРемонта,
                  |	ВТ_Итог.ID КАК ID,
                  |	ВТ_Итог.ИнвентарныйНомер КАК ИнвентарныйНомер,
                  |	ВТ_Итог.Наименование КАК Наименование,
                  |	ПРЕДСТАВЛЕНИЕ(ВТ_Итог.проф_Марка) КАК Марка,
                  |	ВТ_Итог.ЗаводскойНомер КАК ЗаводскойНомер,
//++ Проф-ИТ, #296, Горетовская М.С., 05.10.2023 - Доработки по задаче #128
                  |	ВТ_Итог.ГосНомер КАК ГосНомер,
//-- Проф-ИТ, #296, Горетовская М.С., 05.10.2023 - Доработки по задаче #128
                  |	ВТ_Итог.ЗаявкаНаРемонт КАК ЗаявкаНаРемонт,
                  |	ВТ_Итог.ДатаОкончания КАК Дата,
                  |	ВТ_Итог.ДефектОписание КАК Неисправность,
                  |	ПРЕДСТАВЛЕНИЕ(ВТ_Итог.ПлановыйГрафикРаботы) КАК РежимРаботы,
                  |	ВТ_Итог.ДефектПричина КАК Причина,
                  |	ВТ_Итог.ВремяРаботы КАК ВремяРаботы,
                  |	ВТ_Итог.ВремяПростоя КАК ВремяПростоя
                  |ИЗ
                  |	ВТ_Итог КАК ВТ_Итог";
   
   Если ЗначениеЗаполнено(ОбъектРемонта) Тогда 
	   ТекстЗапроса = СтрЗаменить(ТекстЗапроса,
								 "&ОбъектРемонта",
								 "торо_ВыявленныеДефектыСписокДефектов.ОбъектРемонта = &ОбъектРемонта");
	   Запрос.УстановитьПараметр("ОбъектРемонта", ОбъектРемонта);  
   Иначе
	   ТекстЗапроса = СтрЗаменить(ТекстЗапроса, "&ОбъектРемонта", "ИСТИНА");
   КонецЕсли;    
   
   Возврат ТекстЗапроса;
   
КонецФункции	

&НаСервере
Функция ПолучитьДокументДляПечати()
   
   Табл = Новый ТабличныйДокумент;
   Табл.ОриентацияСтраницы = ОриентацияСтраницы.Ландшафт;
   
   ЭтаОбработка = РеквизитФормыВЗначение("Отчет");
   Макет = ЭтаОбработка.ПолучитьМакет("МакетПечать");
   ОбластьШапкаДокумента = Макет.ПолучитьОбласть("ОбластьШапкаДокумента"); 
   ОбластьШапкаТаблицы = Макет.ПолучитьОбласть("ОбластьШапкаТаблицы");  
   
   ОбластьШапкаДокумента.Параметры.текДата = Формат(ТекущаяДатаСеанса(), "ДФ='«dd» MMMM гггг'' г.'''");
   ОбластьШапкаДокумента.Параметры.месяц = Формат(ПериодОтчета.ДатаНачала, "ДФ='MMMM'");
   ОбластьШапкаДокумента.Параметры.год = Формат(ПериодОтчета.ДатаНачала, "ДФ=yyyy");

   Табл.Вывести(ОбластьШапкаДокумента); 
	   
   Запрос = Новый Запрос;
   Запрос.УстановитьПараметр("ДатаНачала", НачалоМесяца(ПериодОтчета.ДатаНачала));  
   Запрос.УстановитьПараметр("ДатаОкончания", КонецМесяца(ПериодОтчета.ДатаОкончания));  
   Запрос.Текст = ТекстЗапросаПечатьЖурналУчетаЗаявочныхРемонтовЗаМесяц(ОбъектРемонта, Запрос);	     	   
   РезультатЗапроса = Запрос.ВыполнитьПакет();  
   
   ВыборкаОбъектыРемонта = РезультатЗапроса[5].Выбрать();
   ВыборкаНаработка = РезультатЗапроса[3].Выбрать();
   ВыборкаНоменклатура = РезультатЗапроса[4].Выбрать();
   
   Если ВыборкаОбъектыРемонта.Количество() > 0 Тогда
	   Табл.Вывести(ОбластьШапкаТаблицы); 
   КонецЕсли;
   
   Пока ВыборкаОбъектыРемонта.Следующий() Цикл
	   ОбластьСтрока = Макет.ПолучитьОбласть("ОбластьСтрока");
	   ОбластьСтрока.Параметры.Заполнить(ВыборкаОбъектыРемонта); 
	   ОбластьСтрока.Параметры.Наименование = ВыборкаОбъектыРемонта.Наименование + " " + ВыборкаОбъектыРемонта.Марка; 
	   ОбластьСтрока.Параметры.Пробег = "";
	   ОбластьСтрока.Параметры.Наработка = "";
	   
	   Отбор = Новый Структура;
	   Отбор.Вставить("ОбъектРемонта", ВыборкаОбъектыРемонта.ОбъектРемонта);
	   Отбор.Вставить("ID", ВыборкаОбъектыРемонта.ID);
	   
	   ЗаполнитьПробег(Отбор, ВыборкаНаработка, ОбластьСтрока);	
	   
	   ОбластьСтрока.Параметры.ВремяРаботы = ?(ЗначениеЗаполнено(ВыборкаОбъектыРемонта.ВремяРаботы),
		   Строка(ВыборкаОбъектыРемонта.ВремяРаботы) + ", ч",
		   "");
	   ОбластьСтрока.Параметры.ВремяПростоя = ?(ЗначениеЗаполнено(ВыборкаОбъектыРемонта.ВремяПростоя),
		   Строка(ВыборкаОбъектыРемонта.ВремяПростоя) + ", ч",
		   "");  
	   
	   НоменклатураВыведена = Ложь;
	   
	   МассивСтрокНоменклатуры = ПолучитьМассивСтрокНоменклатуры(Отбор, ВыборкаНоменклатура, ОбластьСтрока, Макет);
					   
	   Если Табл.ПроверитьВывод(МассивСтрокНоменклатуры) = Ложь Тогда
		   Табл.ВывестиГоризонтальныйРазделительСтраниц(); 
		   Табл.Вывести(ОбластьШапкаТаблицы); 
	   КонецЕсли;
	   
	   ВывестиМассивСтрокНоменклатуры(МассивСтрокНоменклатуры, Табл);
				 
	   НоменклатураВыведена = ?(МассивСтрокНоменклатуры.Количество() > 0, Истина, Ложь);  
				   
	   Если Не НоменклатураВыведена
	   И Табл.ПроверитьВывод(ОбластьСтрока) = Ложь Тогда	
		   Табл.ВывестиГоризонтальныйРазделительСтраниц(); 
		   Табл.Вывести(ОбластьШапкаТаблицы);
		   Табл.Вывести(ОбластьСтрока);
	   ИначеЕсли Не НоменклатураВыведена Тогда
		   Табл.Вывести(ОбластьСтрока);			
	   Иначе
		   // Объединим ячейки               
		   ОбъединитьЯчейкиВыведеннойНоменклатуре(Табл, МассивСтрокНоменклатуры.Количество());			
	   КонецЕсли;			
   КонецЦикла;
   
   Возврат Табл;
   
КонецФункции

&НаСервере
Процедура ЗаполнитьПробег(Отбор, ВыборкаНаработка, ОбластьСтрока)
   
   Пока ВыборкаНаработка.НайтиСледующий(Отбор) Цикл
	   ОбластьСтрока.Параметры.Пробег = ОбластьСтрока.Параметры.Пробег + ?(ЗначениеЗаполнено(ВыборкаНаработка.Пробег),
	   									Строка(ВыборкаНаработка.Пробег) + ", км", "");
	   ОбластьСтрока.Параметры.Наработка = ОбластьСтрока.Параметры.Наработка
	   											+ ?(ЗначениеЗаполнено(ВыборкаНаработка.Наработка),
	   													Строка(ВыборкаНаработка.Наработка) + ", ч", "");
	   
   КонецЦикла;
   ВыборкаНаработка.Сбросить();		
   
КонецПроцедуры	

&НаСервере
Функция ПолучитьМассивСтрокНоменклатуры(Отбор, ВыборкаНоменклатура, ОбластьСтрока, Макет)
   
   МассивСтрокНоменклатуры = Новый Массив; 
   
   Пока ВыборкаНоменклатура.НайтиСледующий(Отбор) Цикл 
	   ОбластьСтрокаМассив = Макет.ПолучитьОбласть("ОбластьСтрока");
	   ОбластьСтрокаМассив.Параметры.Заполнить(ОбластьСтрока.Параметры);
	   ОбластьСтрокаМассив.Параметры.Заполнить(ВыборкаНоменклатура);
	   МассивСтрокНоменклатуры.Добавить(ОбластьСтрокаМассив);
   КонецЦикла; 
   ВыборкаНоменклатура.Сбросить();	
   
   Возврат МассивСтрокНоменклатуры;
   
КонецФункции	

&НаСервере
Процедура ВывестиМассивСтрокНоменклатуры(МассивСтрокНоменклатуры, ТабДок)
   
   Для Каждого СтрокаМассив Из МассивСтрокНоменклатуры Цикл
	   ТабДок.Вывести(СтрокаМассив);
   КонецЦикла;	
   
КонецПроцедуры	

&НаСервере
Процедура ОбъединитьЯчейкиВыведеннойНоменклатуре(ТабДок, КоличествоСтрокНоменклатуры)
   
   // Объединим ячейки               
   НачальнаяСтрока = Строка(ТабДок.ВысотаТаблицы - КоличествоСтрокНоменклатуры + 1);
   ВысотаТаблицыСтрокой = Строка(ТабДок.ВысотаТаблицы);
   
//++ Проф-ИТ, #296, Горетовская М.С., 05.10.2023 - Доработки по задаче #128
   //Для Сч = 2 По 13 Цикл 
   //    Если Сч <> 10 И Сч <> 11 Тогда
   Для Сч = 2 По 15 Цикл 
	   Если Сч <> 12 И Сч <> 13 Тогда
//-- Проф-ИТ, #296, Горетовская М.С., 05.10.2023 - Доработки по задаче #128
	   АдресОбласти = СтрШаблон("R%1C%2:R%3C%2", НачальнаяСтрока, Сч,  ВысотаТаблицыСтрокой);
	   ТабДок.Область(АдресОбласти).Объединить();  
	   КонецЕсли;
   КонецЦикла;		
   
КонецПроцедуры	

#КонецОбласти

//-- Проф-ИТ, #128, Башинская А., 25.08.2023