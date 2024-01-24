
//++ Проф-ИТ, #195, Башинская А., 13.10.2023

#Область ОбработчикиКомандФормы

&НаКлиенте
Процедура КомандаЗагрузить(Команда)  
	
	Если Не ЗначениеЗаполнено(Объект.Организация) Тогда
		ПоказатьПредупреждение(, "Заполните организацию.");    
		Возврат;
	КонецЕсли;
	
	РезультатЗагрузки = "";
	КомандаЗагрузитьНаСервере();   
	
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

// Процедура записывает логи о загрузке данных из файла для последующего вывода
// Параметры:
// ТекстСообщения - текст записываемого сообщения
// НомерСтроки - по данной строке идет запись сообщений
// Приоритет - 1 - ошибка, 2 - предупреждение, 3 - общая информация, 4 - итоговое сообщение
//
&НаСервере
Процедура ЗаписатьСообщение(ТекстСообщения, НомерСтроки, Приоритет)
	
	НСтр = ТаблицаСообщенийПоСтрокам.Добавить();
	НСтр.Сообщение = ТекстСообщения;
	НСтр.Номер = НомерСтроки;
	НСтр.Приоритет = Приоритет;
	
КонецПроцедуры

&НаСервере
Процедура ВывестиРезультатыЗагрузки()
	
	ТаблицаСообщенийПоСтрокам.Сортировать("Номер, Приоритет");   
	предСтрока = Неопределено;
	Для Каждого Стр Из ТаблицаСообщенийПоСтрокам Цикл
		Если предСтрока <> Неопределено И Стр.Номер <> предСтрока.Номер Тогда
			РезультатЗагрузки = РезультатЗагрузки + символы.ПС; 
		КонецЕсли; 
		
		РезультатЗагрузки = РезультатЗагрузки + Стр.Сообщение + символы.ПС;   
		
		предСтрока = Стр;
	КонецЦикла;
	
КонецПроцедуры

&НаСервере
Процедура КомандаЗагрузитьНаСервере()
	
	СообщениеОшибки = ""; 
	СообщенияЗагрузки = "";     
	СообщенияПредупреждения = "";
	ЕстьНезаполненныеЗначения = Ложь;
	МассивОбязательныхКолонок = ПолучитьСписокОбязательныхКолонок();
	МассивЗаписейРегистра = Новый Соответствие;
	СтрокиИзФайла.Очистить();
	ТаблицаСообщенийПоСтрокам.Очистить();
	
	Для НомерСтроки = 2 По ТабличныйДокумент.ВысотаТаблицы Цикл 
			
		СтрПоля = "";
		Для Каждого эл Из МассивОбязательныхКолонок Цикл
			Если ТабличныйДокумент.Область(НомерСтроки, эл).Текст = "" Тогда
				СтрПоля = СтрШаблон("%1%2;%3", СтрПоля, ТабличныйДокумент.Область(1, эл).Текст, Символы.ПС);
			КонецЕсли;
		КонецЦикла;
		
		Если ЗначениеЗаполнено(СтрПоля) Тогда
			ЕстьНезаполненныеЗначения = Истина;
			Смещение = 2;         
			СтрПоля = Лев(СтрПоля, СтрДлина(СтрПоля) - Смещение);
			       
			СообщениеОшибки = СтрШаблон("Не заполнены обязательные поля в строке №%1 (%2)", НомерСтроки, СтрПоля);			
			ЗаписатьСообщение(СообщениеОшибки, НомерСтроки, 1);
        КонецЕсли;
		
		НСтр = СтрокиИзФайла.Добавить();
		НСтр.НомерСтроки = НомерСтроки;
		НСтр.ОбъектРемонта										 =  СокрЛП(ТабличныйДокумент.Область(НомерСтроки, 1).Текст);
		НСтр.ВидРемонта											 =  СокрЛП(ТабличныйДокумент.Область(НомерСтроки, 3).Текст);
		НСтр.ТиповойОР											 =  СокрЛП(ТабличныйДокумент.Область(НомерСтроки, 5).Текст);
		НСтр.ГрафикРемонтныхРабот								 =  СокрЛП(ТабличныйДокумент.Область(НомерСтроки, 7).Текст);
		НСтр.НормативныйРемонт 									 =  СокрЛП(ТабличныйДокумент.Область(НомерСтроки, 8).Текст);  	
		НСтр.СпособВыполнения								 	 = ПолучитьСпособВыполнения(
																		СокрЛП(ТабличныйДокумент.Область(НомерСтроки, 10).Текст));;	
		НСтр.СпособПланирования									 =  СокрЛП(ТабличныйДокумент.Область(НомерСтроки, 11).Текст);
		НСтр.СпособУчетаДатыОтсчета								 =  ПолучитьСпособУчетаДатыОтсчета(
																		Сокрлп(ТабличныйДокумент.Область(НомерСтроки, 13).Текст));
		НСтр.НеУчаствуетВПланировании 							 =  СокрЛП(ТабличныйДокумент.Область(НомерСтроки, 14).Текст);
		НСтр.НеИспользоватьВПланированииДатуВводаВЭксплуатацию   =  СокрЛП(ТабличныйДокумент.Область(НомерСтроки, 15).Текст);
		НСтр.НеУчитыватьПриПланированииГрафикРаботыОборудования  =  СокрЛП(ТабличныйДокумент.Область(НомерСтроки, 16).Текст);
		НСтр.НеУчитыватьПриПланированииГрафикРемонтныхРабот 	 =  СокрЛП(ТабличныйДокумент.Область(НомерСтроки, 17).Текст);
		НСтр.НеУчитыватьПродолжительность 						 =  СокрЛП(ТабличныйДокумент.Область(НомерСтроки, 18).Текст);
		НСтр.УчитыватьПростоиОборудования 						 =  СокрЛП(ТабличныйДокумент.Область(НомерСтроки, 19).Текст);
		НСтр.РассчитыватьНаработкуВПростое 						 =  СокрЛП(ТабличныйДокумент.Область(НомерСтроки, 20).Текст);  
		НСтр.РазрешитьСокращатьМежремонтныйПериодПриПланировании =  СокрЛП(ТабличныйДокумент.Область(НомерСтроки, 21).Текст);
		НСтр.ИспользоватьДатыЗамещающихРемонтов 				 =  СокрЛП(ТабличныйДокумент.Область(НомерСтроки, 22).Текст);
		
	КонецЦикла;	
	
	Если ЕстьНезаполненныеЗначения Тогда
		ЗаписатьСообщение("Загрузка полностью остановлена.", 1000000, 4); 
		ЗаполнитьПроверитьДанные(МассивЗаписейРегистра, СообщенияЗагрузки, СообщениеОшибки, СообщенияПредупреждения, Ложь);
	Иначе		
		ЗаполнитьПроверитьДанные(МассивЗаписейРегистра, СообщенияЗагрузки, СообщениеОшибки, СообщенияПредупреждения);			
	КонецЕсли;
	
	ВывестиРезультатыЗагрузки();
	
	Сообщение = Новый СообщениеПользователю;
	Сообщение.Текст = "Загрузка завершена";
	Сообщение.Сообщить();

КонецПроцедуры    

&НаСервере
Функция ПолучитьСпособВыполнения(Текст)
	
	Если Текст = "Хозспособ" Тогда
		Возврат Перечисления.СпособыСтроительства.Хозспособ;
	ИначеЕсли Текст = "Подрядный" Тогда
		Возврат Перечисления.СпособыСтроительства.Подрядный;
	Иначе
		Возврат Неопределено;
	КонецЕсли;
	
КонецФункции
																	
&НаСервере
Процедура ВыполнитьЗаписьВБД(СообщенияЗагрузки, МассивЗаписейРегистра)
	
	НачатьТранзакцию(); 
	
	Попытка
		
		// Запишем регистр
		Для Каждого Эл Из МассивЗаписейРегистра Цикл
			
			Запись = Эл.Значение;
			
			Набор = РегистрыСведений.торо_НормативныеРемонтыОборудования.СоздатьНаборЗаписей();
			Набор.Отбор.ОбъектРемонта.Установить(запись.ОбъектРемонта);
			Набор.Отбор.ВидРемонта.Установить(запись.ВидРемонта);
			Набор.Прочитать();   
			
			Если Набор.Количество() > 0 Тогда           
				
				// Обновим запись
				Запись.Записать(Истина);
				СообщенияЗагрузки = СтрШаблон("Запись строки №%1 обновлена%2", Эл.Ключ, Символы.ПС);
				ЗаписатьСообщение(СообщенияЗагрузки, Эл.Ключ, 3);

			Иначе         
				
				Запись.Записать();   
				
				// Создадим запись в регистре торо_ПредшествующиеНормативныеРемонтыОборудования
				// И торо_НормативныеРемонтыОборудования
				Запись2 = РегистрыСведений.торо_ПредшествующиеНормативныеРемонтыОборудования.СоздатьМенеджерЗаписи();
				Запись2.ВидРемонта = Запись.ВидРемонта;
				Запись2.ОбъектРемонта = Запись.ОбъектРемонта;
				Запись2.ТиповойОР = Запись.ТиповойОР;
				Запись2.ПредшествующийВидРемонта = Запись.ВидРемонта;
				Запись2.Приоритет = Истина; 
				Запись2.Записать();  
								
				СообщенияЗагрузки = СтрШаблон("Запись строки №%1 загружена%2", Эл.Ключ, Символы.ПС);
				ЗаписатьСообщение(СообщенияЗагрузки, Эл.Ключ, 3);

			КонецЕсли;
			
		КонецЦикла;   
		
		ЗафиксироватьТранзакцию(); 
		
	Исключение
		ОтменитьТранзакцию();
		ОписаниеОшибки = ОбработкаОшибок.ПодробноеПредставлениеОшибки(ИнформацияОбОшибке());		
		СообщенияЗагрузки = "Возникли ошибки при записи в базу данных. " + ОписаниеОшибки;   
		ЗаписатьСообщение(СообщенияЗагрузки, 1000000, 4);
	КонецПопытки;
	
КонецПроцедуры

&НаСервере
Процедура ЗаполнитьЗаписьРегистра(Стр, МассивЗаписейРегистра)
	
	// Создадим запись в регистре торо_НормативныеРемонтыОборудования	
	Запись = РегистрыСведений.торо_НормативныеРемонтыОборудования.СоздатьМенеджерЗаписи(); 
	ЗаполнитьЗначенияСвойств(Запись, Стр);
	Запись.Приоритет = Истина;
		
	МассивЗаписейРегистра.Вставить(Стр.НомерСтроки, Запись);
		
КонецПроцедуры

&НаСервере
Функция ПолучитьДанныеДляПроверкиИЗаполнения()
	
	ТЗ = РеквизитФормыВЗначение("СтрокиИзФайла");
	
	Запрос = Новый Запрос;
	Запрос.Текст = "ВЫБРАТЬ
	               |	ВнешнийИсточник.ОбъектРемонта КАК ОбъектРемонта,
	               |	ВнешнийИсточник.ВидРемонта КАК ВидРемонта,
	               |	ВнешнийИсточник.НомерСтроки КАК НомерСтроки,
	               |	ВнешнийИсточник.ТиповойОР КАК ТиповойОР,
	               |	ВнешнийИсточник.НормативныйРемонт КАК НормативныйРемонт,
	               |	ВнешнийИсточник.СпособПланирования КАК СпособПланирования,
	               |	ВнешнийИсточник.ГрафикРемонтныхРабот КАК ГрафикРемонтныхРабот,
	               |	ВнешнийИсточник.СпособВыполнения КАК СпособВыполнения,
	               |	ВнешнийИсточник.НеИспользоватьВПланированииДатуВводаВЭксплуатацию КАК НеИспользоватьВПланированииДатуВводаВЭксплуатацию,
	               |	ВнешнийИсточник.СпособУчетаДатыОтсчета КАК СпособУчетаДатыОтсчета,
	               |	ВнешнийИсточник.НеУчаствуетВПланировании КАК НеУчаствуетВПланировании,
	               |	ВнешнийИсточник.РазрешитьСокращатьМежремонтныйПериодПриПланировании КАК РазрешитьСокращатьМежремонтныйПериодПриПланировании,
	               |	ВнешнийИсточник.НеУчитыватьПриПланированииГрафикРаботыОборудования КАК НеУчитыватьПриПланированииГрафикРаботыОборудования,
	               |	ВнешнийИсточник.НеУчитыватьПриПланированииГрафикРемонтныхРабот КАК НеУчитыватьПриПланированииГрафикРемонтныхРабот,
	               |	ВнешнийИсточник.НеУчитыватьПродолжительность КАК НеУчитыватьПродолжительность,
	               |	ВнешнийИсточник.УчитыватьПростоиОборудования КАК УчитыватьПростоиОборудования,
	               |	ВнешнийИсточник.РассчитыватьНаработкуВПростое КАК РассчитыватьНаработкуВПростое,
	               |	ВнешнийИсточник.ИспользоватьДатыЗамещающихРемонтов КАК ИспользоватьДатыЗамещающихРемонтов
	               |ПОМЕСТИТЬ ВнешнийИсточник
	               |ИЗ
	               |	&ТЗ КАК ВнешнийИсточник
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	ВнешнийИсточник.СпособВыполнения КАК СпособВыполнения,
	               |	ВнешнийИсточник.НеИспользоватьВПланированииДатуВводаВЭксплуатацию КАК НеИспользоватьВПланированииДатуВводаВЭксплуатацию,
	               |	ВнешнийИсточник.СпособУчетаДатыОтсчета КАК СпособУчетаДатыОтсчета,
	               |	ВнешнийИсточник.НеУчаствуетВПланировании КАК НеУчаствуетВПланировании,
	               |	ВнешнийИсточник.РазрешитьСокращатьМежремонтныйПериодПриПланировании КАК РазрешитьСокращатьМежремонтныйПериодПриПланировании,
	               |	ВнешнийИсточник.НеУчитыватьПриПланированииГрафикРаботыОборудования КАК НеУчитыватьПриПланированииГрафикРаботыОборудования,
	               |	ВнешнийИсточник.НеУчитыватьПриПланированииГрафикРемонтныхРабот КАК НеУчитыватьПриПланированииГрафикРемонтныхРабот,
	               |	ВнешнийИсточник.НеУчитыватьПродолжительность КАК НеУчитыватьПродолжительность,
	               |	ВнешнийИсточник.УчитыватьПростоиОборудования КАК УчитыватьПростоиОборудования,
	               |	ВнешнийИсточник.РассчитыватьНаработкуВПростое КАК РассчитыватьНаработкуВПростое,
	               |	ВнешнийИсточник.ИспользоватьДатыЗамещающихРемонтов КАК ИспользоватьДатыЗамещающихРемонтов,
	               |	КОЛИЧЕСТВО(РАЗЛИЧНЫЕ ЕСТЬNULL(торо_ОбъектыРемонта.Ссылка, НЕОПРЕДЕЛЕНО)) КАК ОбъектРемонтаКолВо,
	               |	МАКСИМУМ(ЕСТЬNULL(торо_ОбъектыРемонта.Ссылка, НЕОПРЕДЕЛЕНО)) КАК ОбъектРемонта,
	               |	КОЛИЧЕСТВО(РАЗЛИЧНЫЕ ЕСТЬNULL(торо_ВидыРемонтов.Ссылка, НЕОПРЕДЕЛЕНО)) КАК ВидРемонтаКолВо,
	               |	МАКСИМУМ(ЕСТЬNULL(торо_ВидыРемонтов.Ссылка, НЕОПРЕДЕЛЕНО)) КАК ВидРемонта,
	               |	КОЛИЧЕСТВО(РАЗЛИЧНЫЕ ЕСТЬNULL(торо_ТиповыеОР.Ссылка, НЕОПРЕДЕЛЕНО)) КАК ТиповойОРКолВо,
	               |	МАКСИМУМ(ЕСТЬNULL(торо_ТиповыеОР.Ссылка, НЕОПРЕДЕЛЕНО)) КАК ТиповойОР,
	               |	КОЛИЧЕСТВО(РАЗЛИЧНЫЕ ЕСТЬNULL(Календари.Ссылка, НЕОПРЕДЕЛЕНО)) КАК ГрафикРемонтныхРаботКолВо,
	               |	МАКСИМУМ(ЕСТЬNULL(Календари.Ссылка, НЕОПРЕДЕЛЕНО)) КАК ГрафикРемонтныхРабот,
	               |	КОЛИЧЕСТВО(РАЗЛИЧНЫЕ ЕСТЬNULL(торо_ИдентификаторыТехКарт.Ссылка, НЕОПРЕДЕЛЕНО)) КАК НормативныйРемонтКолВо,
	               |	МАКСИМУМ(ЕСТЬNULL(торо_ИдентификаторыТехКарт.Ссылка, НЕОПРЕДЕЛЕНО)) КАК НормативныйРемонт,
	               |	КОЛИЧЕСТВО(РАЗЛИЧНЫЕ ЕСТЬNULL(торо_СпособыПланированияРемонтныхРабот.Ссылка, НЕОПРЕДЕЛЕНО)) КАК СпособПланированияКолВо,
	               |	МАКСИМУМ(ЕСТЬNULL(торо_СпособыПланированияРемонтныхРабот.Ссылка, НЕОПРЕДЕЛЕНО)) КАК СпособПланирования,
	               |	ВнешнийИсточник.НомерСтроки КАК НомерСтроки
	               |ИЗ
	               |	ВнешнийИсточник КАК ВнешнийИсточник
	               |		ЛЕВОЕ СОЕДИНЕНИЕ Справочник.торо_ОбъектыРемонта КАК торо_ОбъектыРемонта
	               |		ПО ВнешнийИсточник.ОбъектРемонта = торо_ОбъектыРемонта.Код
	               |		ЛЕВОЕ СОЕДИНЕНИЕ Справочник.торо_ВидыРемонтов КАК торо_ВидыРемонтов
	               |		ПО ВнешнийИсточник.ВидРемонта = торо_ВидыРемонтов.Код
	               |		ЛЕВОЕ СОЕДИНЕНИЕ Справочник.торо_ТиповыеОР КАК торо_ТиповыеОР
	               |		ПО ВнешнийИсточник.ТиповойОР = торо_ТиповыеОР.Код
	               |		ЛЕВОЕ СОЕДИНЕНИЕ Справочник.Календари КАК Календари
	               |		ПО ВнешнийИсточник.ГрафикРемонтныхРабот = Календари.Наименование
	               |		ЛЕВОЕ СОЕДИНЕНИЕ Справочник.торо_ИдентификаторыТехКарт КАК торо_ИдентификаторыТехКарт
	               |		ПО ВнешнийИсточник.НормативныйРемонт = торо_ИдентификаторыТехКарт.Код
	               |		ЛЕВОЕ СОЕДИНЕНИЕ Справочник.торо_СпособыПланированияРемонтныхРабот КАК торо_СпособыПланированияРемонтныхРабот
	               |		ПО ВнешнийИсточник.СпособПланирования = торо_СпособыПланированияРемонтныхРабот.Код
	               |
	               |СГРУППИРОВАТЬ ПО
	               |	ВнешнийИсточник.НеУчитыватьПриПланированииГрафикРемонтныхРабот,
	               |	ВнешнийИсточник.СпособВыполнения,
	               |	ВнешнийИсточник.НеУчитыватьПродолжительность,
	               |	ВнешнийИсточник.НомерСтроки,
	               |	ВнешнийИсточник.НеИспользоватьВПланированииДатуВводаВЭксплуатацию,
	               |	ВнешнийИсточник.СпособУчетаДатыОтсчета,
	               |	ВнешнийИсточник.НеУчитыватьПриПланированииГрафикРаботыОборудования,
	               |	ВнешнийИсточник.ИспользоватьДатыЗамещающихРемонтов,
	               |	ВнешнийИсточник.РассчитыватьНаработкуВПростое,
	               |	ВнешнийИсточник.РазрешитьСокращатьМежремонтныйПериодПриПланировании,
	               |	ВнешнийИсточник.НеУчаствуетВПланировании,
	               |	ВнешнийИсточник.УчитыватьПростоиОборудования
	               |
	               |УПОРЯДОЧИТЬ ПО
	               |	НомерСтроки УБЫВ";
	
	Запрос.УстановитьПараметр("ТЗ", ТЗ);			   
	РезультатЗапроса = Запрос.Выполнить().Выбрать();
	
	Возврат РезультатЗапроса;
	
КонецФункции

&НаСервере
Функция ПолучитьСписокОбязательныхКолонок()
	
	М = Новый Массив();
	М.Добавить(1); 
	М.Добавить(3);
	М.Добавить(7);
	М.Добавить(8);
	М.Добавить(10);
	М.Добавить(11);
	М.Добавить(13);
	М.Добавить(14);
	М.Добавить(15);
	М.Добавить(16);
	М.Добавить(17);
	М.Добавить(18);
	М.Добавить(19);
	М.Добавить(20);
	М.Добавить(21);
	М.Добавить(22); 
	
	Возврат М;
	
КонецФункции

&НаСервере
Функция ПолучитьСпособУчетаДатыОтсчета(Наименование)
	
	Если Наименование = "По последнему ремонту или дате планирования" Тогда
		Возврат Перечисления.торо_СпособыУчетаДатыОтсчетаПриПланированииРемонтныхРабот.ПоПоследнемуРемонтуИлиДатеПланирования;
		
	ИначеЕсли Наименование = "По дате ввода в эксплуатацию или последнему ремонту" Тогда
		Возврат Перечисления.торо_СпособыУчетаДатыОтсчетаПриПланированииРемонтныхРабот.ИспользоватьДатуВводаВЭксплуатацию;
		
	ИначеЕсли Наименование = "Использовать дату планирования" Тогда
		Возврат Перечисления.торо_СпособыУчетаДатыОтсчетаПриПланированииРемонтныхРабот.ИспользоватьДатуПланированияВКачествеДатыОтсчета;
		
	ИначеЕсли Наименование = "По последнему ремонту" Тогда
		Возврат Перечисления.торо_СпособыУчетаДатыОтсчетаПриПланированииРемонтныхРабот.НеИспользоватьДатуВводаВЭксплуатацию;
		
	ИначеЕсли Наименование = "По дате ввода в эксплуатацию или плановым датам" Тогда
		Возврат Перечисления.торо_СпособыУчетаДатыОтсчетаПриПланированииРемонтныхРабот.ПоДатеВводаВЭксплуатациюИлиПлановымДатам;   
		
	Иначе 
		Возврат Неопределено;
	КонецЕсли;
	
КонецФункции

&НаСервере
Процедура ЗаполнитьПроверитьДанные(МассивЗаписейРегистра, СообщенияЗагрузки, СообщениеОшибки, СообщенияПредупреждения,
								   ЗаписьВозможна = Истина)

	ДополненныеДанные = ПолучитьДанныеДляПроверкиИЗаполнения();
	ЕстьНенайденныеСсылкиПоКраснымПолям = Ложь;
	// Обработка данных и формирование массива записей регистра
	
	Пока ДополненныеДанные.Следующий() Цикл
		
		Стр = ДополненныеДанные;
						
		ПроверитьДополнительноДанные(Стр, СообщенияПредупреждения, ЕстьНенайденныеСсылкиПоКраснымПолям);
		Если ЕстьОшибкиПоОбязательнымИзмерениям(Стр, СообщениеОшибки) Тогда
			Продолжить;
		КонецЕсли;      
		
		ЗаполнитьЗаписьРегистра(Стр, МассивЗаписейРегистра);
		
	КонецЦикла;
	
	Если ЕстьНенайденныеСсылкиПоКраснымПолям Тогда 
		Если ЗаписьВозможна Тогда // Ранее не записывали это сообщение
			ЗаписатьСообщение("Загрузка полностью остановлена.", 1000000, 4);      
		КонецЕсли;
		Возврат;
	КонецЕсли;     
	
	Если ЗаписьВозможна Тогда
		ВыполнитьЗаписьВБД(СообщенияЗагрузки, МассивЗаписейРегистра); 
	КонецЕсли;

КонецПроцедуры	

&НаСервере
Процедура ПроверитьДополнительноДанные(Стр, СообщенияПредупреждения, ЕстьНенайденныеСсылкиПоКраснымПолям)

	Если Стр.ТиповойОРКолВо > 1 Тогда				
		СообщенияПредупреждения = СтрШаблон("В строке №%1 найдено более одного значения" +
			" для поля шаблона ""Код типового ОР""", Стр.НомерСтроки);   
		ЗаписатьСообщение(СообщенияПредупреждения, Стр.НомерСтроки, 2); 
	КонецЕсли;  
	
	Если Стр.ТиповойОР = Неопределено Тогда				
		СообщенияПредупреждения = СтрШаблон("В строке №%1 не найдено значение" +
			" поля шаблона ""Код типового ОР""", Стр.НомерСтроки); 
		ЗаписатьСообщение(СообщенияПредупреждения, Стр.НомерСтроки, 2);  
	КонецЕсли; 

	Если Стр.ГрафикРемонтныхРаботКолВо > 1 Тогда				
		СообщенияПредупреждения = СтрШаблон("В строке №%1 найдено более одного значения" +
			" для поля шаблона ""График ремонтных работ""", Стр.НомерСтроки);  
		ЗаписатьСообщение(СообщенияПредупреждения, Стр.НомерСтроки, 2);
	КонецЕсли; 
	
	Если Стр.ГрафикРемонтныхРабот = Неопределено Тогда				
		СообщенияПредупреждения = СтрШаблон("В строке №%1 не найдено значение" +
			" поля шаблона ""График ремонтных работ""", Стр.НомерСтроки); 
		ЗаписатьСообщение(СообщенияПредупреждения, Стр.НомерСтроки, 2);    
		ЕстьНенайденныеСсылкиПоКраснымПолям = Истина;
	КонецЕсли; 
	
	Если Стр.НормативныйРемонтКолВо > 1 Тогда				
		СообщенияПредупреждения = СтрШаблон("В строке №%1 найдено более одного значения" +
			" для поля шаблона ""Код технологической карты""", Стр.НомерСтроки);   
		ЗаписатьСообщение(СообщенияПредупреждения, Стр.НомерСтроки, 2);  
	КонецЕсли;   
	
	Если Стр.НормативныйРемонт = Неопределено Тогда				
		СообщенияПредупреждения = СтрШаблон("В строке №%1 не найдено значение" +
			" поля шаблона ""Код технологической карты""", Стр.НомерСтроки);   
		ЗаписатьСообщение(СообщенияПредупреждения, Стр.НомерСтроки, 2);  
		ЕстьНенайденныеСсылкиПоКраснымПолям = Истина;
	КонецЕсли;
	
	Если Стр.СпособПланированияКолВо > 1 Тогда				
		СообщенияПредупреждения = СтрШаблон("В строке №%1 найдено более одного значения" +
			" для поля шаблона ""Код способа планирования""", Стр.НомерСтроки); 
		ЗаписатьСообщение(СообщенияПредупреждения, Стр.НомерСтроки, 2);
	КонецЕсли;
	
	Если Стр.СпособПланирования = Неопределено Тогда				
		СообщенияПредупреждения = СтрШаблон("В строке №%1 не найдено значение" +
			" поля шаблона ""Код способа планирования""", Стр.НомерСтроки); 
		ЗаписатьСообщение(СообщенияПредупреждения, Стр.НомерСтроки, 2);
		ЕстьНенайденныеСсылкиПоКраснымПолям = Истина;
	КонецЕсли;
	
	Если Не ЗначениеЗаполнено(Стр.СпособВыполнения) Тогда				
		СообщенияПредупреждения = СтрШаблон("В строке №%1 не найдено значение" +
			" поля шаблона ""Способ выполнения""", Стр.НомерСтроки); 
		ЗаписатьСообщение(СообщенияПредупреждения, Стр.НомерСтроки, 2);  
		ЕстьНенайденныеСсылкиПоКраснымПолям = Истина;
	КонецЕсли;
	
	Если Не ЗначениеЗаполнено(Стр.СпособУчетаДатыОтсчета) Тогда				
		СообщенияПредупреждения = СтрШаблон("В строке №%1 не найдено значение" +
			" поля шаблона ""Способ учета даты отсчета""", Стр.НомерСтроки); 
		ЗаписатьСообщение(СообщенияПредупреждения, Стр.НомерСтроки, 2);   
		ЕстьНенайденныеСсылкиПоКраснымПолям = Истина;
	КонецЕсли;	

КонецПроцедуры	

&НаСервере
Функция ЕстьОшибкиПоОбязательнымИзмерениям(Стр, СообщениеОшибки)
	
	ЕстьОшибкаПоОбязательнымИзмерениям = Ложь;
	
	Если Стр.ОбъектРемонта = Неопределено Тогда				
		СообщениеОшибки = СтрШаблон("В строке №%1 не найдено значение" +
		" поля шаблона ""Код ОР""", Стр.НомерСтроки);  
		ЗаписатьСообщение(СообщениеОшибки, Стр.НомерСтроки, 1);
		ЕстьОшибкаПоОбязательнымИзмерениям = Истина;
	КонецЕсли; 
	Если Стр.ОбъектРемонтаКолВо > 1 Тогда				
		СообщениеОшибки = СтрШаблон("В строке №%1 найдено более одного значения" +
		" поля для шаблона ""Код ОР""", Стр.НомерСтроки);
		ЗаписатьСообщение(СообщениеОшибки, Стр.НомерСтроки, 1);
		ЕстьОшибкаПоОбязательнымИзмерениям = Истина;
	КонецЕсли; 
	
	Если Стр.ВидРемонта = Неопределено Тогда				
		СообщениеОшибки = СтрШаблон("В строке №%1 не найдено значение" +
		" поля шаблона ""Виды ремонта""", Стр.НомерСтроки);  
		ЗаписатьСообщение(СообщениеОшибки, Стр.НомерСтроки, 1);
		ЕстьОшибкаПоОбязательнымИзмерениям = Истина;
	КонецЕсли;	
	Если Стр.ВидРемонтаКолВо > 1 Тогда				
		СообщениеОшибки = СтрШаблон("В строке №%1 найдено более одного значения" +
		" для поля шаблона ""Виды ремонта""", Стр.НомерСтроки);  
		ЗаписатьСообщение(СообщениеОшибки, Стр.НомерСтроки, 1);
		ЕстьОшибкаПоОбязательнымИзмерениям = Истина;
	КонецЕсли;	
	
	Возврат ЕстьОшибкаПоОбязательнымИзмерениям;
	
КонецФункции	

#КонецОбласти

//-- Проф-ИТ, #195, Башинская А., 13.10.2023
