
#Область ОбработчикиСобытий

Процедура ПриКомпоновкеРезультата(ДокументРезультат, ДанныеРасшифровки, СтандартнаяОбработка)
	
	СтандартнаяОбработка = Ложь;
	
	ДокументРезультат.Очистить();
	Макет = ПолучитьМакет("МакетСводнойСтраницы");
	Настройки = КомпоновщикНастроек.ПолучитьНастройки();

	КлючВарианта = КомпоновщикНастроек.ПользовательскиеНастройки.ДополнительныеСвойства.КлючВарианта;
	
	ПараметрыОтчета = Настройки.ПараметрыДанных.Элементы;
	
	СписокВариантовАнализа = ПолучитьЗначениеПараметраИзНастроек(ПараметрыОтчета, "ПоказателиKPI", Истина);
	
	СтруктураПараметров = Новый Структура;
	СтруктураПараметров.Вставить("СписокВариантовАнализа", СписокВариантовАнализа);
	
	ДозаполнитьСтруктуруПользовательскихНастроек(ПараметрыОтчета, СтруктураПараметров);
	ДозаполнитьСтруктуруДаннымиФормы(СтруктураПараметров);
	
	торо_ПоказателиKPI.ПолучитьДанныеПоказателей(СтруктураПараметров);
	ТаблицаВариантовАнализа = торо_ПоказателиKPI.ПолучитьСписокВариантовАнализаСУчетомРазделителя(СтруктураПараметров.ДанныеПоказателей, СписокВариантовАнализа);
	
	НетДоступныхПоказателей = (СписокВариантовАнализа.Количество()=0);
	торо_ПоказателиKPI.ВывестиСообщениеНетДоступныхПоказателей(ДокументРезультат, Макет, НетДоступныхПоказателей);
	
	КомпоновщикНастроек.ПользовательскиеНастройки.ДополнительныеСвойства.Вставить("ОтчетПустой", НетДоступныхПоказателей); 
	
	СтруктураПараметров.ЦелевыеЗначенияПоказателей = торо_ПоказателиKPI.ПолучитьЦелевыеЗначенияПоказателей(СтруктураПараметров);
	
	Если КлючВарианта = "Основной" Тогда
		
		// Вывод заголовка и параметров
		торо_ПоказателиKPI.ВывестиЗаголовок(ДокументРезультат, Макет);
		торо_ПоказателиKPI.ВывестиПараметры(ДокументРезультат, Макет, СтруктураПараметров);
		торо_ПоказателиKPI.ВывестиДатуВремяФормирования(ДокументРезультат, Макет);
		
		// Вывод показателей
		КоличествоКолонок = ПолучитьЗначениеПараметраИзНастроек(ПараметрыОтчета, "КоличествоКолонокПоказателя");
		Если КоличествоКолонок = 0 Тогда
			КоличествоКолонок = 2;
		КонецЕсли;
		
		КоличествоСтрок = торо_ПоказателиKPIКлиентСервер.ОкруглитьВверх(ТаблицаВариантовАнализа.Количество() / КоличествоКолонок);
		
		СчетчикЭлементовСтроки = 1;
		СчетчикСтрок = 1;
		
		НомерЯчейки = 0;
		
		Для Каждого ЭлементСписка Из ТаблицаВариантовАнализа Цикл 
			НомерЯчейки = НомерЯчейки + 1;
			ВариантАнализа = ЭлементСписка.ВариантАнализа;
			Показатель = ЭлементСписка.Показатель;
			Разделитель = ЭлементСписка.Разделитель;
			
			ВывестиПоказательНаПечать(ДокументРезультат, Макет, Показатель, ВариантАнализа, Разделитель, НомерЯчейки, СтруктураПараметров);
			
			СчетчикЭлементовСтроки = СчетчикЭлементовСтроки + 1;
			Если СчетчикЭлементовСтроки > КоличествоКолонок Тогда 
				СчетчикЭлементовСтроки = 1;
				СчетчикСтрок = СчетчикСтрок + 1;
				
				// Для последней строки отступ выводить не требуется
				Если СчетчикСтрок <= КоличествоСтрок Тогда
					торо_ПоказателиKPI.ВывестиОтступ(ДокументРезультат, Макет);
				КонецЕсли;
			КонецЕсли;
		КонецЦикла;
		
		торо_ПоказателиKPI.ВывестиОтступ(ДокументРезультат, Макет);
		
		НомерЯчейки = 0;
		Макет = ПолучитьМакет("МакетПодробнойСтраницыПоказателя");
		
		Для Каждого ЭлементСписка Из СписокВариантовАнализа Цикл
			НомерЯчейки = НомерЯчейки + 1;
			
			ВариантАнализа = ЭлементСписка.ВариантАнализа;
			Показатель = ЭлементСписка.Показатель;
			Разделитель = ЭлементСписка.Разделитель;
			Если НЕ ВариантАнализа.ВариантОтображенияПоУмолчанию = Перечисления.торо_ВариантыОтображенияВариантовАнализа.Кратко
				И НЕ ВариантАнализа.ТипДиаграммы = Перечисления.торо_ТипыДиаграммПоказателей.Измерительная Тогда
				
				ВывестиПоказательНаПечать(ДокументРезультат, Макет, Показатель, ВариантАнализа, Разделитель, НомерЯчейки, СтруктураПараметров, Истина);
				торо_ПоказателиKPI.ВывестиОтступ(ДокументРезультат, Макет);
				
			КонецЕсли;

		КонецЦикла;
		
	Иначе
		
		МакетИнтегральнойОценки = ПолучитьМакет("МакетИнтегральнойОценки");
		
		// Вывод заголовка и параметров
		торо_ПоказателиKPI.ВывестиЗаголовок(ДокументРезультат, МакетИнтегральнойОценки);
		торо_ПоказателиKPI.ВывестиПараметры(ДокументРезультат, МакетИнтегральнойОценки, СтруктураПараметров);
		торо_ПоказателиKPI.ВывестиДатуВремяФормирования(ДокументРезультат, МакетИнтегральнойОценки);
		
		// Вывод показателей
		СвернутаяТабВариантовАнализа = торо_ПоказателиKPI.ПолучитьСписокПоказателейДляИнтегральнойОценки(СписокВариантовАнализа);
		Если СвернутаяТабВариантовАнализа.Количество() < 3 Тогда
			КомпоновщикНастроек.ПользовательскиеНастройки.ДополнительныеСвойства.Вставить("ОтчетПустой", Истина);
			Возврат;
		КонецЕсли;
		
		торо_ПоказателиKPI.ЗаполнитьИнтегральнуюОценку(СтруктураПараметров.ИнтегральнаяДиаграмма, СвернутаяТабВариантовАнализа, СтруктураПараметров);	
		торо_ПоказателиKPI.ВывестиИнтегральнуюОценкуНаПечать(ДокументРезультат, МакетИнтегральнойОценки, СтруктураПараметров.ИнтегральнаяДиаграмма);
		
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

#Область ФормированиеОтчета

Процедура ВывестиПоказательНаПечать(ДокументРезультат, Макет, Показатель, ВариантАнализа, Разделитель, НомерЯчейки, СтруктураПараметров, Подробно = Ложь)
	
	ВариантОтображения = ВариантАнализа.ВариантОтображенияПоУмолчанию;
	
	СтрокаПоказателя = СтруктураПараметров.ДанныеПоказателей.НайтиСтроки(Новый Структура("Показатель, Разделитель", Показатель, Разделитель));
	Если СтрокаПоказателя.Количество() = 0 Тогда
		Возврат;
	КонецЕсли;
	СтрокаПоказателя = СтрокаПоказателя[0];
	
	Если ЭтоАдресВременногоХранилища(СтрокаПоказателя.АдресХранилищаДанных) Тогда
		ТабДанныхИсх = ПолучитьИзВременногоХранилища(СтрокаПоказателя.АдресХранилищаДанных);
		ТабДанных = ТабДанныхИсх.Скопировать();
	Иначе
		Возврат;
	КонецЕсли;
	
	Если ЭтоАдресВременногоХранилища(СтрокаПоказателя.АдресХранилищаДанныхСПериодами) Тогда
		ТабДанныхСПериодамиИсх = ПолучитьИзВременногоХранилища(СтрокаПоказателя.АдресХранилищаДанныхСПериодами);
		ТабДанныхСПериодами = ТабДанныхСПериодамиИсх.Скопировать();
	Иначе
		Возврат;
	КонецЕсли;
	
	Попытка
		СтруктураДанныхДляВывода = торо_ПоказателиKPI.ПолучитьСтруктураДанныхДляВыводаВМонитор(ТабДанных, ТабДанныхСПериодами, ВариантАнализа, СтруктураПараметров);
	Исключение
		ОписаниеОшибки = ОписаниеОшибки();
		ШаблонОшибки = НСтр("ru='При обработке данных показателя ""%1"" произошла ошибка: %2'");
		ТекстСообщения = СтрШаблон(ШаблонОшибки, ВариантАнализа, ОписаниеОшибки);
		#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда
			ОбщегоНазначения.СообщитьПользователю(ТекстСообщения);
		#Иначе
		   ОбщегоНазначенияКлиент.СообщитьПользователю(ТекстСообщения);
		#КонецЕсли
		Возврат;
	КонецПопытки;
	
	Если СтруктураДанныхДляВывода = Неопределено Тогда
		Возврат;
	КонецЕсли;
	
	ВременныйДокумент = Новый ТабличныйДокумент;
	торо_ПоказателиKPI.ВывестиЗаголовокПоказателяНаПечать(ВременныйДокумент, Макет, Показатель, ВариантАнализа, Разделитель);
	
	Если Подробно Тогда
		
		Если ВариантОтображения = Перечисления.торо_ВариантыОтображенияВариантовАнализа.Диаграмма Тогда
			ВывестиПодробноТаблицуДиаграммыНаПечать(ВременныйДокумент, Макет, ВариантАнализа, СтруктураДанныхДляВывода, СтруктураПараметров);
			торо_ПоказателиKPI.ВывестиОтступ(ВременныйДокумент, Макет);
			ВывестиПоказательНаПечатьДиаграмма(ВременныйДокумент, Макет, ВариантАнализа, СтруктураДанныхДляВывода, СтруктураПараметров);
		ИначеЕсли ВариантОтображения = Перечисления.торо_ВариантыОтображенияВариантовАнализа.Таблица Тогда
			ВывестиПоказательНаПечатьТаблица(ВременныйДокумент, Макет, ВариантАнализа, СтруктураДанныхДляВывода, СтруктураПараметров, Подробно);
		КонецЕсли;
		
	Иначе
	
		Если ВариантОтображения = Перечисления.торо_ВариантыОтображенияВариантовАнализа.Диаграмма Тогда
			ВывестиПоказательНаПечатьДиаграмма(ВременныйДокумент, Макет, ВариантАнализа, СтруктураДанныхДляВывода, СтруктураПараметров);
		ИначеЕсли ВариантОтображения = Перечисления.торо_ВариантыОтображенияВариантовАнализа.Кратко Тогда
			ВывестиПоказательНаПечатьКратко(ВременныйДокумент, Макет, ВариантАнализа, СтруктураДанныхДляВывода, СтруктураПараметров);
		ИначеЕсли ВариантОтображения = Перечисления.торо_ВариантыОтображенияВариантовАнализа.Таблица Тогда
			ВывестиПоказательНаПечатьТаблица(ВременныйДокумент, Макет, ВариантАнализа, СтруктураДанныхДляВывода, СтруктураПараметров);
		КонецЕсли;
	
	КонецЕсли;
	
	ДокументРезультат.Присоединить(ВременныйДокумент);
	
КонецПроцедуры

Процедура ВывестиПоказательНаПечатьДиаграмма(ДокументРезультат, Макет, ВариантАнализа, СтруктураДанныхДляВывода, СтруктураПараметров)
	
	ДиаграммаИсточник = Новый Диаграмма;
	ДиаграммаИсточник.ТипДиаграммы = торо_ПоказателиKPIКлиентСервер.ПолучитьТипДиаграммыПоЗначениюПеречисления(ВариантАнализа.ТипДиаграммы);
	
	торо_ПоказателиKPI.ВывестиДанныеПоказателяВДиаграмму(СтруктураДанныхДляВывода, ВариантАнализа, ДиаграммаИсточник, СтруктураПараметров);
	торо_ПоказателиKPI.ВывестиПоказательНаПечатьДиаграмма(ДокументРезультат, Макет, ДиаграммаИсточник);
	
КонецПроцедуры

Процедура ВывестиПоказательНаПечатьКратко(ДокументРезультат, Макет, ВариантАнализа, СтруктураДанныхДляВывода, СтруктураПараметров)
	
	СтрокаРезультат = "";
	
	торо_ПоказателиKPI.ВывестиДанныеПоказателяВКраткомВиде(СтруктураДанныхДляВывода, ВариантАнализа, СтрокаРезультат);
	торо_ПоказателиKPI.ВывестиПоказательНаПечатьКратко(ДокументРезультат, Макет, СтрокаРезультат);
		
КонецПроцедуры

Процедура ВывестиПоказательНаПечатьТаблица(ДокументРезультат, Макет, ВариантАнализа, СтруктураДанныхДляВывода, СтруктураПараметров, Подробно = Ложь)
	
	ТаблицаРезультат = Новый ТаблицаЗначений;
	ТаблицаРезультат.Колонки.Добавить("Период", Новый ОписаниеТипов("Дата"));
	ТаблицаРезультат.Колонки.Добавить("Разрез", Новый ОписаниеТипов("Строка"));
	ТаблицаРезультат.Колонки.Добавить("Значение", Новый ОписаниеТипов("Строка"));
	
	торо_ПоказателиKPI.ВывестиДанныеПоказателяВТаблицу(СтруктураДанныхДляВывода, ВариантАнализа, ТаблицаРезультат);
	
	Если Подробно Тогда
		
		торо_ПоказателиKPI.ВывестиПодробноПоказательНаПечатьТаблица(ДокументРезультат, Макет, ВариантАнализа, СтруктураПараметров, ТаблицаРезультат);
		
	Иначе
	
		торо_ПоказателиKPI.ВывестиПоказательНаПечатьТаблица(ДокументРезультат, Макет, ВариантАнализа, СтруктураПараметров, ТаблицаРезультат);
	
	КонецЕсли;
	
КонецПроцедуры

Процедура ВывестиПодробноТаблицуДиаграммыНаПечать(ДокументРезультат, Макет, ВариантАнализа, СтруктураДанныхДляВывода, СтруктураПараметров)

	ТаблицаРезультат = Новый ТаблицаЗначений;
	ТаблицаРезультат.Колонки.Добавить("Период", Новый ОписаниеТипов("Дата"));
	ТаблицаРезультат.Колонки.Добавить("Разрез", Новый ОписаниеТипов("Строка"));
	ТаблицаРезультат.Колонки.Добавить("Значение", Новый ОписаниеТипов("Строка"));
	
	ДиаграммаИсточник = Новый Диаграмма;
	ДиаграммаИсточник.ТипДиаграммы = торо_ПоказателиKPIКлиентСервер.ПолучитьТипДиаграммыПоЗначениюПеречисления(ВариантАнализа.ТипДиаграммы);
	
	торо_ПоказателиKPI.ВывестиДанныеПоказателяВДиаграмму(СтруктураДанныхДляВывода, ВариантАнализа, ДиаграммаИсточник, СтруктураПараметров);
	
	Для Каждого СерияИсточник из ДиаграммаИсточник.Серии Цикл
		
		НоваяСтрока = ТаблицаРезультат.Добавить();
		НоваяСтрока.Период = ТекущаяДата();
		НоваяСтрока.Разрез = СерияИсточник.Текст;
		
		Для Каждого ТочкаИсточник из ДиаграммаИсточник.Точки Цикл
			
			НоваяСтрока.Период = ТочкаИсточник.Текст;
			
			Попытка 
				ЗначениеИсточник = ДиаграммаИсточник.ПолучитьЗначение(ТочкаИсточник, СерияИсточник);
				НоваяСтрока.Значение = ЗначениеИсточник.Значение;
			Исключение
			КонецПопытки;
			
		КонецЦикла;
		
	КонецЦикла;
	
	торо_ПоказателиKPI.ВывестиПодробноПоказательНаПечатьТаблица(ДокументРезультат, Макет, ВариантАнализа, СтруктураПараметров, ТаблицаРезультат);

КонецПроцедуры

#КонецОбласти

Функция ПолучитьЗначениеПараметраИзНастроек(ПараметрыОтчета, ИмяПараметра, ПреобразоватьКСпискуЗначений = Ложь)
	
	ЗначениеПараметра = ПараметрыОтчета.Найти(ИмяПараметра);
	Если ЗначениеПараметра <> Неопределено Тогда
		ЗначениеПараметра = ЗначениеПараметра.Значение;
	КонецЕсли;
	
	ЗначениеДляВозврата = ЗначениеПараметра;
	
	Если ПреобразоватьКСпискуЗначений И ТипЗнч(ЗначениеПараметра) <> Тип("СписокЗначений") Тогда
		ЗначениеДляВозврата = Новый СписокЗначений;
		Если ЗначениеПараметра <> Неопределено Тогда
			ЗначениеДляВозврата.Добавить(ЗначениеПараметра);
		КонецЕсли;
	КонецЕсли;
	
	Возврат ЗначениеДляВозврата;
	
КонецФункции

Функция ПолучитьИспользованиеПараметраИзПользовательскихНастроек(ПараметрыОтчета, ИмяПараметра)
	
	ИспользованиеПараметра = Ложь;
	
	Для каждого ЭлементНастройки из ПараметрыОтчета Цикл
		Если Строка(ЭлементНастройки.Параметр) = ИмяПараметра Тогда
			ИспользованиеПараметра = ЭлементНастройки.Использование;
			Прервать;
		КонецЕсли;
	КонецЦикла;
	
	Возврат ИспользованиеПараметра;
	
КонецФункции


Процедура ДозаполнитьСтруктуруПользовательскихНастроек(ПараметрыОтчета, СтруктураПараметров)
	
	// Получим настройки и отборы
	СтруктураПараметров.Вставить("ОтборПоПериоду", ПолучитьЗначениеПараметраИзНастроек(ПараметрыОтчета, "Период"));
	СтруктураПараметров.Вставить("ОтборПоПериодуСравнения", ПолучитьЗначениеПараметраИзНастроек(ПараметрыОтчета, "ПериодСравнения"));
	СтруктураПараметров.Вставить("ТипПоказателей", ПолучитьЗначениеПараметраИзНастроек(ПараметрыОтчета, "ТипПоказателей"));
	СтруктураПараметров.Вставить("ТипСерийИнтегральнойОценки", ПолучитьЗначениеПараметраИзНастроек(ПараметрыОтчета, "ТипСерийИнтегральнойОценки"));
	
	ИспользоватьДетализациюПоПериоду = ПолучитьИспользованиеПараметраИзПользовательскихНастроек(ПараметрыОтчета, "ДетализацияПоПериоду");
	Если ИспользоватьДетализациюПоПериоду Тогда
		ДетализацияПоПериоду = ПолучитьЗначениеПараметраИзНастроек(ПараметрыОтчета, "ДетализацияПоПериоду");
	Иначе
		ДетализацияПоПериоду = Неопределено;
	КонецЕсли;
	
	СтруктураПараметров.Вставить("ИспользоватьДетализациюПоПериоду", ИспользоватьДетализациюПоПериоду);
	СтруктураПараметров.Вставить("ДетализацияПоПериоду", ДетализацияПоПериоду);
	
	ИспользоватьОтборПоОрганизации = ПолучитьИспользованиеПараметраИзПользовательскихНастроек(ПараметрыОтчета, "Организация");
	Если ИспользоватьОтборПоОрганизации Тогда
		ОтборПоОрганизации = ПолучитьЗначениеПараметраИзНастроек(ПараметрыОтчета, "Организация", Истина);
	Иначе
		ОтборПоОрганизации = Новый СписокЗначений;
	КонецЕсли;
	
	СтруктураПараметров.Вставить("ИспользоватьОтборПоОрганизации", ИспользоватьОтборПоОрганизации);
	СтруктураПараметров.Вставить("ОтборПоОрганизации", ОтборПоОрганизации);
	
	ИспользоватьОтборПоПодразделению = ПолучитьИспользованиеПараметраИзПользовательскихНастроек(ПараметрыОтчета, "Подразделение");
	Если ИспользоватьОтборПоПодразделению Тогда
		ОтборПоПодразделению = ПолучитьЗначениеПараметраИзНастроек(ПараметрыОтчета, "Подразделение", Истина);
	Иначе
		ОтборПоПодразделению = Новый СписокЗначений;
	КонецЕсли;

	СтруктураПараметров.Вставить("ИспользоватьОтборПоПодразделению", ИспользоватьОтборПоПодразделению);
	СтруктураПараметров.Вставить("ОтборПоПодразделению", ОтборПоПодразделению);
	
	ИспользоватьОтборПоОР = ПолучитьИспользованиеПараметраИзПользовательскихНастроек(ПараметрыОтчета, "ОбъектРемонта");
	Если ИспользоватьОтборПоОР Тогда
		ОтборПоОР = ПолучитьЗначениеПараметраИзНастроек(ПараметрыОтчета, "ОбъектРемонта", Истина);
	Иначе
		ОтборПоОР = Новый СписокЗначений;
	КонецЕсли;
	
	СтруктураПараметров.Вставить("ИспользоватьОтборПоОР", ИспользоватьОтборПоОР);
	СтруктураПараметров.Вставить("ОтборПоОР", ОтборПоОР);
		
КонецПроцедуры

Процедура ДозаполнитьСтруктуруДаннымиФормы(СтруктураПараметров)
	
	ОписаниеТиповРазделителя = Новый ОписаниеТипов("Строка, Булево, Число, Дата");
	ОписаниеТиповРазделителя = Новый ОписаниеТипов(ОписаниеТиповРазделителя, Справочники.ТипВсеСсылки().Типы());
	ОписаниеТиповРазделителя = Новый ОписаниеТипов(ОписаниеТиповРазделителя, Документы.ТипВсеСсылки().Типы());
	ОписаниеТиповРазделителя = Новый ОписаниеТипов(ОписаниеТиповРазделителя, Перечисления.ТипВсеСсылки().Типы());
	ОписаниеТиповРазделителя = Новый ОписаниеТипов(ОписаниеТиповРазделителя, ПланыСчетов.ТипВсеСсылки().Типы());
	ОписаниеТиповРазделителя = Новый ОписаниеТипов(ОписаниеТиповРазделителя, ПланыВидовХарактеристик.ТипВсеСсылки().Типы());
	ОписаниеТиповРазделителя = Новый ОписаниеТипов(ОписаниеТиповРазделителя, Задачи.ТипВсеСсылки().Типы());
	ОписаниеТиповРазделителя = Новый ОписаниеТипов(ОписаниеТиповРазделителя, ПланыВидовРасчета.ТипВсеСсылки().Типы());
	ОписаниеТиповРазделителя = Новый ОписаниеТипов(ОписаниеТиповРазделителя, ПланыОбмена.ТипВсеСсылки().Типы());
	ОписаниеТиповРазделителя = Новый ОписаниеТипов(ОписаниеТиповРазделителя, БизнесПроцессы.ТипВсеСсылки().Типы());
	
	СтруктураПараметров.Вставить("УникальныйИдентификатор", Новый УникальныйИдентификатор);

	ДанныеПоказателей = Новый ТаблицаЗначений;
	ДанныеПоказателей.Колонки.Добавить("Показатель", Новый ОписаниеТипов("СправочникСсылка.торо_ПоказателиKPI"));
	ДанныеПоказателей.Колонки.Добавить("АдресХранилищаДанных", Новый ОписаниеТипов("Строка"));
	ДанныеПоказателей.Колонки.Добавить("ВариантЗаполнения", Новый ОписаниеТипов("ПеречислениеСсылка.торо_ВариантыЗаполненияПоказателяKPI"));
	ДанныеПоказателей.Колонки.Добавить("ИдентификаторДляФормул", Новый ОписаниеТипов("Строка"));
	ДанныеПоказателей.Колонки.Добавить("ТекстФормулы", Новый ОписаниеТипов("Строка"));
	ДанныеПоказателей.Колонки.Добавить("ТипПоказателя", Новый ОписаниеТипов("ПеречислениеСсылка.торо_ТипыПоказателейKPI"));
	ДанныеПоказателей.Колонки.Добавить("АдресХранилищаДанныхСПериодами", Новый ОписаниеТипов("Строка"));
	ДанныеПоказателей.Колонки.Добавить("Разделитель", ОписаниеТиповРазделителя);

	ДанныеПоказателейЗаПрошлыйПериод = ДанныеПоказателей.Скопировать();
	
	ТаблицаДат = Новый ТаблицаЗначений;
	ТаблицаДат.Колонки.Добавить("Период", Новый ОписаниеТипов("Дата"));
	ТаблицаДат.Колонки.Добавить("НачалоПериода", Новый ОписаниеТипов("Дата"));
	ТаблицаДат.Колонки.Добавить("КонецПериода", Новый ОписаниеТипов("Дата"));
	
	ТаблицаДатЗаПрошлыйПериод = ТаблицаДат.Скопировать();
	ТаблицаДатНаТекДату = ТаблицаДат.Скопировать();
	ТаблицаДатНаТекДатуЗаПрошлыйПериод = ТаблицаДат.Скопировать();
	
	ЦелевыеЗначенияПоказателей = Новый ТаблицаЗначений;
	ЦелевыеЗначенияПоказателей.Колонки.Добавить("Показатель", Новый ОписаниеТипов("СправочникСсылка.торо_ПоказателиKPI"));
	ЦелевыеЗначенияПоказателей.Колонки.Добавить("ЦелевоеЗначение", Новый ОписаниеТипов("Число"));
	ЦелевыеЗначенияПоказателей.Колонки.Добавить("КритическоеЗначение", Новый ОписаниеТипов("Число"));
	ЦелевыеЗначенияПоказателей.Колонки.Добавить("Организация", Новый ОписаниеТипов("СправочникСсылка.Организации"));
	ЦелевыеЗначенияПоказателей.Колонки.Добавить("Подразделение", Новый ОписаниеТипов("СправочникСсылка.СтруктураПредприятия"));
	ЦелевыеЗначенияПоказателей.Колонки.Добавить("ОбъектРемонта", Новый ОписаниеТипов("СправочникСсылка.торо_ОбъектыРемонта"));
	
	СтруктураПараметров.Вставить("ДанныеПоказателей", ДанныеПоказателей);
	СтруктураПараметров.Вставить("ДанныеПоказателейЗаПрошлыйПериод", ДанныеПоказателейЗаПрошлыйПериод);

	ИнтегральнаяДиаграмма = Новый Диаграмма;
	ИнтегральнаяДиаграмма.ТипДиаграммы = ТипДиаграммы.РадарныйГрафик;
	СтруктураПараметров.Вставить("ИнтегральнаяДиаграмма", ИнтегральнаяДиаграмма);
	
	СтруктураПараметров.Вставить("ТаблицаДат", ТаблицаДат);
	СтруктураПараметров.Вставить("ТаблицаДатЗаПрошлыйПериод", ТаблицаДатЗаПрошлыйПериод);
	СтруктураПараметров.Вставить("ТаблицаДатНаТекДату", ТаблицаДатНаТекДату);
	СтруктураПараметров.Вставить("ТаблицаДатНаТекДатуЗаПрошлыйПериод", ТаблицаДатНаТекДатуЗаПрошлыйПериод);
	
	СтруктураПараметров.Вставить("ЦелевыеЗначенияПоказателей", ЦелевыеЗначенияПоказателей);

КонецПроцедуры

#КонецОбласти