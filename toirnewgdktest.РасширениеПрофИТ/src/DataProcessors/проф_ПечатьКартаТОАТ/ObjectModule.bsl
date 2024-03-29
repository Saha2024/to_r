
//++ Проф-ИТ, #76, Башинская А., 13.10.2023

#Область ПрограммныйИнтерфейс

Функция ПечатьКартаТОАТ(МассивОбъектов, ПараметрыПечати) Экспорт 
	
	ТабДок = Новый ТабличныйДокумент;
	
	Макет = ПолучитьМакет("проф_КартаТОАТ");
	
	ОбластьПодвал = Макет.ПолучитьОбласть("ОбластьПодвал"); 
	
	Запрос = Новый Запрос;
	Запрос.УстановитьПараметр("МассивОбъектов", МассивОбъектов);
	Запрос.УстановитьПараметр("МассивID", ПараметрыПечати.МассивID);
	Запрос.Текст = ТекстЗапросаПечатьКартаТОАТ(); 
	
	ОбщаяВыборка = Запрос.ВыполнитьПакет(); 
	ТЗОР = ОбщаяВыборка[0].Выбрать(); 
	ВыборкаРемонтныеРаботы = ОбщаяВыборка[1].Выбрать(); 
	ВыборкаЗапЧасти = ОбщаяВыборка[2].Выбрать();  
		
	Пока ТЗОР.Следующий() Цикл	
		Отбор = Новый Структура;
		ОбластьШапки   = Макет.ПолучитьОбласть("ОбластьШапка");
		ЗаполнитьЗначенияСвойств(ОбластьШапки.Параметры, ТЗОР);
		
		НомерНаПечать = ПрефиксацияОбъектовКлиентСервер.НомерНаПечать(ТЗОР.Номер, Истина, Истина);
		ОбластьШапки.Параметры.Номер = НомерНаПечать;

		ТабДок.Вывести(ОбластьШапки);
		
		Отбор.Вставить("ID", ТЗОР.ID);
		
		ВывестиРемонтныеРаботы(ВыборкаРемонтныеРаботы, Отбор, ТабДок, Макет);
		ТабДок.ВывестиГоризонтальныйРазделительСтраниц();
		
		ВывестиСопутствующийРемонт(ТабДок, Макет);  
		
		ВывестиПереченьУстановленныхЗЧ(ВыборкаЗапЧасти, Отбор, ТабДок, Макет);
		
		ТабДок.Вывести(ОбластьПодвал); 
		ТабДок.ВывестиГоризонтальныйРазделительСтраниц();
	КонецЦикла;
	
	Возврат ТабДок;
	
КонецФункции

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

Функция ТекстЗапросаПечатьКартаТОАТ()
	
	ТекстЗапроса = "ВЫБРАТЬ
       |	ЕСТЬNULL(торо_АктОВыполненииЭтапаРабот.Дата, ДАТАВРЕМЯ(1, 1, 1)) КАК Дата,
       |	ЕСТЬNULL(торо_АктОВыполненииЭтапаРабот.Номер, """") КАК Номер,
       |	ЕСТЬNULL(торо_ОбъектыРемонта.проф_Марка, """") КАК Марка,
       |	ЕСТЬNULL(торо_ОбъектыРемонта.проф_Гараж, """") КАК Гараж,
       |	ЕСТЬNULL(торо_ОбъектыРемонта.проф_Госномер, """") КАК Госномер,
       |	торо_АктОВыполненииЭтапаРаботРемонтыОборудования.ID КАК ID,
       |	торо_АктОВыполненииЭтапаРаботРемонтыОборудования.Ссылка КАК Ссылка,
       |	торо_АктОВыполненииЭтапаРаботРемонтыОборудования.НомерСтроки КАК НомерСтроки,
       |	ЕСТЬNULL(торо_АктОВыполненииЭтапаРабот.Подразделение, ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка)) КАК Подразделение
       |ИЗ
       |	Документ.торо_АктОВыполненииЭтапаРабот.РемонтыОборудования КАК торо_АктОВыполненииЭтапаРаботРемонтыОборудования
       |		ЛЕВОЕ СОЕДИНЕНИЕ Документ.торо_АктОВыполненииЭтапаРабот КАК торо_АктОВыполненииЭтапаРабот
       |		ПО торо_АктОВыполненииЭтапаРаботРемонтыОборудования.Ссылка = торо_АктОВыполненииЭтапаРабот.Ссылка
       |		ЛЕВОЕ СОЕДИНЕНИЕ Справочник.торо_ОбъектыРемонта КАК торо_ОбъектыРемонта
       |		ПО торо_АктОВыполненииЭтапаРаботРемонтыОборудования.ОбъектРемонта = торо_ОбъектыРемонта.Ссылка
       |ГДЕ
       |	торо_АктОВыполненииЭтапаРаботРемонтыОборудования.Ссылка В(&МассивОбъектов)
       |	И торо_АктОВыполненииЭтапаРаботРемонтыОборудования.ID В(&МассивID)
       |
       |ОБЪЕДИНИТЬ ВСЕ
       |
       |ВЫБРАТЬ
       |	ЕСТЬNULL(торо_ЗаявкаНаРемонт.Дата, ДАТАВРЕМЯ(1, 1, 1)),
       |	ЕСТЬNULL(торо_ЗаявкаНаРемонт.Номер, """"),
       |	ЕСТЬNULL(торо_ОбъектыРемонта.проф_Марка, """"),
       |	ЕСТЬNULL(торо_ОбъектыРемонта.проф_Гараж, """"),
       |	ЕСТЬNULL(торо_ОбъектыРемонта.проф_Госномер, """"),
       |	торо_ЗаявкаНаРемонтРемонтыОборудования.ID,
       |	торо_ЗаявкаНаРемонтРемонтыОборудования.Ссылка,
       |	торо_ЗаявкаНаРемонтРемонтыОборудования.НомерСтроки,
       |	ЕСТЬNULL(торо_ЗаявкаНаРемонт.Подразделение, ЗНАЧЕНИЕ(Справочник.СтруктураПредприятия.ПустаяСсылка))
       |ИЗ
       |	Документ.торо_ЗаявкаНаРемонт.РемонтыОборудования КАК торо_ЗаявкаНаРемонтРемонтыОборудования
       |		ЛЕВОЕ СОЕДИНЕНИЕ Документ.торо_ЗаявкаНаРемонт КАК торо_ЗаявкаНаРемонт
       |		ПО торо_ЗаявкаНаРемонтРемонтыОборудования.Ссылка = торо_ЗаявкаНаРемонт.Ссылка
       |		ЛЕВОЕ СОЕДИНЕНИЕ Справочник.торо_ОбъектыРемонта КАК торо_ОбъектыРемонта
       |		ПО торо_ЗаявкаНаРемонтРемонтыОборудования.ОбъектРемонта = торо_ОбъектыРемонта.Ссылка
       |ГДЕ
       |	торо_ЗаявкаНаРемонтРемонтыОборудования.Ссылка В(&МассивОбъектов)
       |	И торо_ЗаявкаНаРемонтРемонтыОборудования.ID В(&МассивID)
       |
       |УПОРЯДОЧИТЬ ПО
       |	НомерСтроки
       |;
       |
       |////////////////////////////////////////////////////////////////////////////////
       |ВЫБРАТЬ
       |	торо_АктОВыполненииЭтапаРаботРемонтныеРаботы.Ссылка КАК Ссылка,
       |	торо_АктОВыполненииЭтапаРаботРемонтныеРаботы.РемонтыОборудования_ID КАК ID,
       |	ПРЕДСТАВЛЕНИЕ(торо_АктОВыполненииЭтапаРаботРемонтныеРаботы.РемонтнаяРабота) КАК РемонтнаяРабота,
       |	торо_АктОВыполненииЭтапаРаботРемонтныеРаботы.Родитель_ID КАК Родитель_ID,
       |	торо_АктОВыполненииЭтапаРаботРемонтныеРаботы.Выполнено КАК Выполнено,
       |	ЕСТЬNULL(торо_ТехнологическиеОперации.СодержаниеРабот, """") КАК СодержаниеРабот,
       |	торо_АктОВыполненииЭтапаРаботРемонтныеРаботы.НомерСтроки КАК НомерСтроки
       |ИЗ
       |	Документ.торо_АктОВыполненииЭтапаРабот.РемонтныеРаботы КАК торо_АктОВыполненииЭтапаРаботРемонтныеРаботы
       |		ЛЕВОЕ СОЕДИНЕНИЕ Справочник.торо_ТехнологическиеОперации КАК торо_ТехнологическиеОперации
       |		ПО торо_АктОВыполненииЭтапаРаботРемонтныеРаботы.РемонтнаяРабота = торо_ТехнологическиеОперации.Ссылка
       |ГДЕ
       |	торо_АктОВыполненииЭтапаРаботРемонтныеРаботы.Ссылка В(&МассивОбъектов)
       |
       |ОБЪЕДИНИТЬ ВСЕ
       |
       |ВЫБРАТЬ
       |	торо_ЗаявкаНаРемонтРемонтныеРаботы.Ссылка,
       |	торо_ЗаявкаНаРемонтРемонтныеРаботы.РемонтыОборудования_ID,
       |	ПРЕДСТАВЛЕНИЕ(торо_ЗаявкаНаРемонтРемонтныеРаботы.РемонтнаяРабота),
       |	торо_ЗаявкаНаРемонтРемонтныеРаботы.Родитель_ID,
       |	ЛОЖЬ,
       |	ЕСТЬNULL(торо_ТехнологическиеОперации.СодержаниеРабот, """"),
       |	торо_ЗаявкаНаРемонтРемонтныеРаботы.НомерСтроки
       |ИЗ
       |	Документ.торо_ЗаявкаНаРемонт.РемонтныеРаботы КАК торо_ЗаявкаНаРемонтРемонтныеРаботы
       |		ЛЕВОЕ СОЕДИНЕНИЕ Справочник.торо_ТехнологическиеОперации КАК торо_ТехнологическиеОперации
       |		ПО торо_ЗаявкаНаРемонтРемонтныеРаботы.РемонтнаяРабота = торо_ТехнологическиеОперации.Ссылка
       |ГДЕ
       |	торо_ЗаявкаНаРемонтРемонтныеРаботы.Ссылка В(&МассивОбъектов)
       |
       |УПОРЯДОЧИТЬ ПО
       |	торо_АктОВыполненииЭтапаРаботРемонтныеРаботы.НомерСтроки
       |;
       |
       |////////////////////////////////////////////////////////////////////////////////
       |ВЫБРАТЬ
       |	торо_АктОВыполненииЭтапаРаботСерийныеЗапчасти.Ссылка КАК Ссылка,
       |	торо_АктОВыполненииЭтапаРаботСерийныеЗапчасти.РемонтыОборудования_ID КАК ID,
       |	ПРЕДСТАВЛЕНИЕ(торо_АктОВыполненииЭтапаРаботСерийныеЗапчасти.Номенклатура) КАК Номенклатура,
       |	ПРЕДСТАВЛЕНИЕ(торо_АктОВыполненииЭтапаРаботСерийныеЗапчасти.ХарактеристикаНоменклатуры) КАК ХарактеристикаНоменклатуры,
       |	торо_АктОВыполненииЭтапаРаботСерийныеЗапчасти.НомерСтроки КАК НомерСтроки
       |ИЗ
       |	Документ.торо_АктОВыполненииЭтапаРабот.СерийныеЗапчасти КАК торо_АктОВыполненииЭтапаРаботСерийныеЗапчасти
       |ГДЕ
       |	торо_АктОВыполненииЭтапаРаботСерийныеЗапчасти.Ссылка В(&МассивОбъектов)
       |
       |ОБЪЕДИНИТЬ ВСЕ
       |
       |ВЫБРАТЬ
       |	торо_ЗаявкаНаРемонтЗапчастиРемонта.Ссылка,
       |	торо_ЗаявкаНаРемонтЗапчастиРемонта.РемонтыОборудования_ID,
       |	ПРЕДСТАВЛЕНИЕ(торо_ЗаявкаНаРемонтЗапчастиРемонта.Номенклатура),
       |	ПРЕДСТАВЛЕНИЕ(торо_ЗаявкаНаРемонтЗапчастиРемонта.ХарактеристикаНоменклатуры),
       |	торо_ЗаявкаНаРемонтЗапчастиРемонта.НомерСтроки
       |ИЗ
       |	Документ.торо_ЗаявкаНаРемонт.ЗапчастиРемонта КАК торо_ЗаявкаНаРемонтЗапчастиРемонта
       |ГДЕ
       |	торо_ЗаявкаНаРемонтЗапчастиРемонта.Ссылка В(&МассивОбъектов)
       |
       |УПОРЯДОЧИТЬ ПО
       |	НомерСтроки";

	Возврат ТекстЗапроса;
	
КонецФункции	

Процедура ВывестиРемонтныеРаботы(ВыборкаРемонтныеРаботы, Отбор, ТабДок, Макет)
	
	ОбластьШапкаРаботы  = Макет.ПолучитьОбласть("ОбластьШапкаРаботы");
	ТабДок.Вывести(ОбластьШапкаРаботы);
	
	Пока ВыборкаРемонтныеРаботы.НайтиСледующий(Отбор) Цикл 
		ОбластьСтрокаРодитель  = Макет.ПолучитьОбласть("ОбластьСтрокаРаботыРодитель"); 
		ОбластьСтрокаРабота  = Макет.ПолучитьОбласть("ОбластьСтрокаРаботы"); 	
		Если ЗначениеЗаполнено(ВыборкаРемонтныеРаботы.Родитель_ID) Тогда 
			ОбластьСтрокаРабота.Параметры.Работа = ВыборкаРемонтныеРаботы.РемонтнаяРабота;
			ОбластьСтрокаРабота.Параметры.Выполнено = ?(ВыборкаРемонтныеРаботы.Выполнено = Истина, "b", "");
			ТабДок.Вывести(ОбластьСтрокаРабота);
			Если ЗначениеЗаполнено(ВыборкаРемонтныеРаботы.СодержаниеРабот) Тогда
				ОбластьСтрокаТехОперация  = Макет.ПолучитьОбласть("ОбластьСтрокаРаботы");
				ОбластьСтрокаТехОперация.Параметры.Работа = ВыборкаРемонтныеРаботы.СодержаниеРабот; 
				ТабДок.Вывести(ОбластьСтрокаТехОперация);
			КонецЕсли;
		Иначе
			ОбластьСтрокаРодитель.Параметры.Родитель = ВыборкаРемонтныеРаботы.РемонтнаяРабота;
			ОбластьСтрокаРодитель.Параметры.Выполнено = ?(ВыборкаРемонтныеРаботы.Выполнено = Истина, "b", "");
			ТабДок.Вывести(ОбластьСтрокаРодитель);
		КонецЕсли;	
	КонецЦикла;	
	ВыборкаРемонтныеРаботы.Сбросить();		
	
КонецПроцедуры	

Процедура ВывестиСопутствующийРемонт(ТабДок, Макет)
	
	ОбластьШапкаСопутРемонт = Макет.ПолучитьОбласть("ОбластьШапкаСопутРемонт");
	ТабДок.Вывести(ОбластьШапкаСопутРемонт);
	
	Для сч = 1 По 11 Цикл 
		ОбластьСтрокаСопутРемонт = Макет.ПолучитьОбласть("ОбластьСтрокаСопутРемонт"); 
		ОбластьСтрокаСопутРемонт.Параметры.НомерПП = сч; 
		ТабДок.Вывести(ОбластьСтрокаСопутРемонт);
	КонецЦикла;	
		
КонецПроцедуры

Процедура ВывестиПереченьУстановленныхЗЧ(ВыборкаЗапЧасти, Отбор, ТабДок, Макет)
	
	ОбластьШапкаЗЧ  = Макет.ПолучитьОбласть("ОбластьШапкаЗЧ");
	ТабДок.Вывести(ОбластьШапкаЗЧ);
	сч = 0;
	
	Пока ВыборкаЗапЧасти.НайтиСледующий(Отбор) Цикл 
		ОбластьСтрокаЗЧ  = Макет.ПолучитьОбласть("ОбластьСтрокаЗЧ"); 
		сч = сч + 1;
		ОбластьСтрокаЗЧ.Параметры.НомерПП = Сч;
		ОбластьСтрокаЗЧ.Параметры.Наименование = СтрШаблон("%1 %2",
			ВыборкаЗапЧасти.Номенклатура,
			ВыборкаЗапЧасти.ХарактеристикаНоменклатуры);				
		ТабДок.Вывести(ОбластьСтрокаЗЧ);
	КонецЦикла;	
	ВыборкаЗапЧасти.Сбросить();		
	
КонецПроцедуры

#КонецОбласти

//-- Проф-ИТ, #76, Башинская А., 13.10.2023