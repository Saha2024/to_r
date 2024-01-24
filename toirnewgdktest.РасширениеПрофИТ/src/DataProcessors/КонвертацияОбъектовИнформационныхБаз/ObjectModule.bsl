
#Область ОписаниеТиповКоллекций

//++ Проф-ИТ, #352, Корнилов М.С., 11.12.2023
&ИзменениеИКонтроль("КоллекцияПравилаВыгрузкиДанных")
Функция проф_КоллекцияПравилаВыгрузкиДанных()
    #Удаление
	Возврат ТаблицаПравилВыгрузки;
    #КонецУдаления
	#Вставка
	ТаблицаПравилВыгрузки.Сортировать("Порядок Возр");
	
	Возврат ТаблицаПравилВыгрузки;
	#КонецВставки
КонецФункции
//-- Проф-ИТ, #352, Корнилов М.С., 11.12.2023

#КонецОбласти

#Область Прочее

//++ Проф-ИТ, #352, Корнилов М.С., 11.12.2023
&Вместо("ВыполнитьВыгрузкуЗарегистрированныхДанных")
Процедура проф_ВыполнитьВыгрузкуЗарегистрированныхДанных(ЗаписьСообщения, СтрокаСообщенияОбОшибке, ТаблицаПравилВыгрузкиИспользуемые)
	
	// Переменные-заглушки для поддержки механизма отладки кода обработчиков событий.
	Перем Отказ, ИмяПКО, ВыборкаДанных, ИсходящиеДанные;
	// {ОБРАБОТЧИК ПередПолучениемИзмененныхОбъектов}
	Если Не ПустаяСтрока(Конвертация.ПередПолучениемИзмененныхОбъектов) Тогда
		
		Попытка
			
			Получатель = УзелДляОбмена;
			
			Если ОтладкаОбработчиковВыгрузки Тогда
				
				ВыполнитьОбработчик_Конвертация_ПередПолучениемИзмененныхОбъектов(Получатель, УзелДляФоновогоОбмена);
				
			Иначе
				
				Выполнить(Конвертация.ПередПолучениемИзмененныхОбъектов);
				
			КонецЕсли;
			
		Исключение
			ЗаписатьИнформациюОбОшибкеОбработчикиКонвертации(175, ПодробноеПредставлениеОшибки(ИнформацияОбОшибке()), НСтр("ru = 'ПередПолучениемИзмененныхОбъектов (конвертация)'"));
			Возврат;
		КонецПопытки;
		
	КонецЕсли;
	// {ОБРАБОТЧИК ПередПолучениемИзмененныхОбъектов}
	
	МассивВыгружаемыхМетаданных = ТаблицаПравилВыгрузкиИспользуемые.ВыгрузитьКолонку("ОбъектВыборкиМетаданные");
	
	// Значение "Неопределено" свидетельствует о необходимости выгрузки констант.
	Если МассивВыгружаемыхМетаданных.Найти(Неопределено) <> Неопределено Тогда
		
		ДополнитьМассивВыгружаемыхМетаданныхКонстантами(МассивВыгружаемыхМетаданных);
		
	КонецЕсли;
	
	// Удаляем из массива элементы со значением "Неопределено".
	УдалитьНедопустимыеЗначенияИзМассиваВыгружаемыхМетаданных(МассивВыгружаемыхМетаданных);
	
	// Регистр сведений СоответствияОбъектовИнформационныхБаз выгружается отдельно, поэтому в эту выборку он не должен
	// входить.
	Если МассивВыгружаемыхМетаданных.Найти(Метаданные.РегистрыСведений.СоответствияОбъектовИнформационныхБаз) <> Неопределено Тогда
		
		ОбщегоНазначенияКлиентСервер.УдалитьЗначениеИзМассива(МассивВыгружаемыхМетаданных, Метаданные.РегистрыСведений.СоответствияОбъектовИнформационныхБаз);
		
	КонецЕсли;
	
	// Обновляем повторно используемые значения МРО.
	ОбменДаннымиСлужебный.ПроверитьКэшМеханизмаРегистрацииОбъектов();
	
	НачальнаяВыгрузкаДанных = ОбменДаннымиСервер.УстановленПризнакНачальнойВыгрузкиДанных(ЗаписьСообщения.Получатель);
	
	МэпИзменений = Новый Соответствие;
	МассивИзменений = Новый Массив;
	
	// ВЫБОРКА ИЗМЕНЕНИЙ
	ВыборкаИзменений = ОбменДаннымиСервер.ВыбратьИзменения(ЗаписьСообщения.Получатель, ЗаписьСообщения.НомерСообщения, МассивВыгружаемыхМетаданных);
	
	ОбъектМетаданныхПредыдущий      = Неопределено;
	ПравилоВыгрузкиДанныхПредыдущее = Неопределено;
	ПравилоВыгрузкиДанных           = Неопределено;
	ВыгружаетсяРегистр              = Ложь;
	ВыгружаютсяКонстанты            = Ложь;
	
	ЭтоОбменЧерезВнешнееСоединение = ЭтоОбменЧерезВнешнееСоединение();
	
	НачатьТранзакциюПриЗагрузкеДанных();
	Попытка 
		УзелДляОбменаОбъект = УзелДляОбмена.ПолучитьОбъект();
		
		Пока ВыборкаИзменений.Следующий() Цикл
			Данные = ВыборкаИзменений.Получить();
			Ключ = МассивВыгружаемыхМетаданных.Найти(Данные.Метаданные());
			Если Ключ = Неопределено Тогда 
				ВызватьИсключение "Не найден ключ в массиве выгружаемых метаданных";
			КонецЕсли;

			МассивДанных = МэпИзменений.Получить(Ключ);
			Если МассивДанных = Неопределено Тогда
				МассивДанных = Новый Массив;
				МассивДанных.Добавить(Данные);				
			Иначе
				МассивДанных.Добавить(Данные);
			КонецЕсли;
			МэпИзменений.Вставить(Ключ, МассивДанных);
			
			Инкремент(КоличествоОбъектовКВыгрузке);
		КонецЦикла;
		
		Для Сч = 0 По МассивВыгружаемыхМетаданных.Количество() -1 Цикл 
			МассивДанных = МэпИзменений.Получить(Сч);
			Если МассивДанных <> Неопределено Тогда
				Для Каждого Элемент Из МассивДанных Цикл 
					МассивИзменений.Добавить(Элемент);
				КонецЦикла;
			КонецЕсли;
		КонецЦикла;                                  
		
		//Пока ВыборкаИзменений.Следующий() Цикл
		//	Инкремент(КоличествоОбъектовКВыгрузке);
		//КонецЦикла;
		
		ВыборкаИзменений.Сбросить();
		
		ЕстьПолеСчетчикЗагруженныхОбъектовВнешнееСоединение = Ложь;
		Если ЭтоОбменЧерезВнешнееСоединение() Тогда
			РеквизитыОбработкиЗагрузки = ОбработкаДляЗагрузкиДанных().Метаданные().Реквизиты;
			// В базе-корреспонденте может не быть этого реквизита обработки.
			Если РеквизитыОбработкиЗагрузки.Найти("КоличествоОбъектовКЗагрузкеВнешнееСоединение") <> Неопределено Тогда
				ОбработкаДляЗагрузкиДанных().КоличествоОбъектовКЗагрузкеВнешнееСоединение = КоличествоОбъектовКВыгрузке;
			КонецЕсли;
			ЕстьПолеСчетчикЗагруженныхОбъектовВнешнееСоединение = РеквизитыОбработкиЗагрузки.Найти("СчетчикЗагруженныхОбъектовВнешнееСоединение") <> Неопределено;
		КонецЕсли;
		
		Для Каждого Данные Из МассивИзменений Цикл
		//Пока ВыборкаИзменений.Следующий() Цикл
			
			Если ЭтоОбменЧерезВнешнееСоединение() Тогда
				// В базе-корреспонденте может не быть этого реквизита обработки.
				Если ЕстьПолеСчетчикЗагруженныхОбъектовВнешнееСоединение Тогда
					Инкремент(ОбработкаДляЗагрузкиДанных().СчетчикЗагруженныхОбъектовВнешнееСоединение);
				КонецЕсли;
			КонецЕсли;
			Инкремент(ПолеСчетчикВыгруженныхОбъектов);
			
			ОбменДаннымиСервер.РассчитатьПроцентВыгрузки(СчетчикВыгруженныхОбъектов(), КоличествоОбъектовКВыгрузке);
			
			//Данные = ВыборкаИзменений.Получить();
			
			ТипДанныхДляВыгрузки = ТипЗнч(Данные);
			
			// Отрабатываем удаление объекта.
			Если ТипДанныхДляВыгрузки = ТипУдалениеОбъекта Тогда
				
				ОтработатьУдалениеОбъекта(Данные);
				Продолжить;
				
			ИначеЕсли ТипДанныхДляВыгрузки = ТипРегистрСоответствия Тогда
				Продолжить;
			КонецЕсли;
			
			ОбъектМетаданныхТекущий = Данные.Метаданные(); // ОбъектМетаданных
			
			// Выгружается новый тип объекта метаданных.
			Если ОбъектМетаданныхПредыдущий <> ОбъектМетаданныхТекущий Тогда
				
				Если ОбъектМетаданныхПредыдущий <> Неопределено Тогда
					
					// {ОБРАБОТЧИК ПослеОбработки ПВД}
					Если ПравилоВыгрузкиДанныхПредыдущее <> Неопределено
						И Не ПустаяСтрока(ПравилоВыгрузкиДанныхПредыдущее.ПослеОбработки) Тогда
						
						Попытка
							
							Если ОтладкаОбработчиковВыгрузки Тогда
								
								ВыполнитьОбработчик_ПВД_ПослеОбработкиПравила(ИмяПКО, ПравилоВыгрузкиДанныхПредыдущее, ИсходящиеДанные);
								
							Иначе
								
								Выполнить(ПравилоВыгрузкиДанныхПредыдущее.ПослеОбработки);
								
							КонецЕсли;
							
						Исключение
							ЗаписатьИнформациюОбОшибкеОбработчикиПВД(32, ПодробноеПредставлениеОшибки(ИнформацияОбОшибке()), ПравилоВыгрузкиДанныхПредыдущее["Имя"], "ПослеОбработкиВыгрузкиДанных");
						КонецПопытки;
						
					КонецЕсли;
					// {ОБРАБОТЧИК ПослеОбработки ПВД}
					
				КонецЕсли;
				
				ОбъектМетаданныхПредыдущий = ОбъектМетаданныхТекущий;
				
				ВыгружаетсяРегистр = Ложь;
				ВыгружаютсяКонстанты = Ложь;
				
				СтруктураДанных = МенеджерыДляПлановОбмена[ОбъектМетаданныхТекущий];
				
				Если СтруктураДанных = Неопределено Тогда
					
					ВыгружаютсяКонстанты = Метаданные.Константы.Содержит(ОбъектМетаданныхТекущий);
					
				ИначеЕсли СтруктураДанных.ЭтоРегистр = Истина Тогда
					
					ВыгружаетсяРегистр = Истина;
					
				КонецЕсли;
				
				Если ВыгружаютсяКонстанты Тогда
					
					ПравилоВыгрузкиДанных = ТаблицаПравилВыгрузкиИспользуемые.Найти(Тип("КонстантыНабор"), "ОбъектВыборкиМетаданные");
					Если ПравилоВыгрузкиДанных = Неопределено Тогда
						 
						 ПравилоВыгрузкиДанных = ТаблицаПравилВыгрузкиИспользуемые.Найти(Тип("КонстантыНабор"), "ОбъектВыборки");
						 
					КонецЕсли;
					
				Иначе
					
					ПравилоВыгрузкиДанных = ТаблицаПравилВыгрузкиИспользуемые.Найти(ОбъектМетаданныхТекущий, "ОбъектВыборкиМетаданные");
					
				КонецЕсли;
				
				ПравилоВыгрузкиДанныхПредыдущее = ПравилоВыгрузкиДанных;
				
				// {ОБРАБОТЧИК ПередОбработкой ПВД}
				ИсходящиеДанные = Неопределено;
				
				Если ПравилоВыгрузкиДанных <> Неопределено
					И Не ПустаяСтрока(ПравилоВыгрузкиДанных.ПередОбработкой) Тогда
					
					Попытка
						
						Если ОтладкаОбработчиковВыгрузки Тогда
							
							ВыполнитьОбработчик_ПВД_ПередОбработкойПравила(Отказ, ИмяПКО, ПравилоВыгрузкиДанных, ИсходящиеДанные, ВыборкаДанных);
							
						Иначе
							
							Выполнить(ПравилоВыгрузкиДанных.ПередОбработкой);
							
						КонецЕсли;
						
					Исключение
						ЗаписатьИнформациюОбОшибкеОбработчикиПВД(31, ПодробноеПредставлениеОшибки(ИнформацияОбОшибке()), ПравилоВыгрузкиДанных.Имя, "ПередОбработкойВыгрузкиДанных");
					КонецПопытки;
					
					
				КонецЕсли;
				// {ОБРАБОТЧИК ПередОбработкой ПВД}
				
			КонецЕсли;
			
			Если ТипДанныхДляВыгрузки <> ТипРегистрСоответствия Тогда
				
				// Определяем вид отправки объекта.
				ОтправкаЭлемента = ОтправкаЭлементаДанных.Авто;
				
				СтандартныеПодсистемыСервер.ПриОтправкеДанныхПодчиненному(Данные, ОтправкаЭлемента, НачальнаяВыгрузкаДанных, УзелДляОбменаОбъект);
				
				Если ОтправкаЭлемента = ОтправкаЭлементаДанных.Удалить Тогда
					
					Если ВыгружаетсяРегистр Тогда
						
						// Удаление регистра отсылаем в виде пустого набора записей.
						
					Иначе
						
						// Посылаем информацию об удалении.
						ОтработатьУдалениеОбъекта(Данные);
						Продолжить;
						
					КонецЕсли;
					
				ИначеЕсли ОтправкаЭлемента = ОтправкаЭлементаДанных.Игнорировать Тогда
					
					Продолжить;
					
				КонецЕсли;
				
			КонецЕсли;
			
			// ВЫГРУЗКА ОБЪЕКТА
			Если ВыгружаетсяРегистр Тогда
				
				// выгрузка регистра
				ВыгрузкаРегистра(Данные, ПравилоВыгрузкиДанных, ИсходящиеДанные, НеВыгружатьОбъектыПоСсылкам);
				
			ИначеЕсли ВыгружаютсяКонстанты Тогда
				
				// выгрузка набора констант
				Свойства = Менеджеры[Тип("КонстантыНабор")];
				
				ВыгрузитьНаборКонстант(ПравилоВыгрузкиДанных, Свойства, ИсходящиеДанные, ОбъектМетаданныхТекущий.Имя);
				
			Иначе
				
				Попытка
					
					СсылкаДляПереходаПриВозникновенииОшибки = Данные.Ссылка;
					
				Исключение
					
					Если Ложь Тогда // Для работы АПК
						
						ВызватьИсключение;
						
					КонецЕсли;
					
				КонецПопытки;
				
				// выгрузка ссылочных типов
				ВыгрузкаОбъектаВыборки(Данные, ПравилоВыгрузкиДанных, , ИсходящиеДанные, НеВыгружатьОбъектыПоСсылкам);
				
			КонецЕсли;
			
			ПроверитьНачалоИФиксациюТранзакцииПриЗагрузкеДанных();
			
		КонецЦикла;
		
		Если ОбъектМетаданныхПредыдущий <> Неопределено Тогда
			
			// {ОБРАБОТЧИК ПослеОбработки ПВД}
			Если ПравилоВыгрузкиДанных <> Неопределено
				И Не ПустаяСтрока(ПравилоВыгрузкиДанных.ПослеОбработки) Тогда
				
				Попытка
					
					Если ОтладкаОбработчиковВыгрузки Тогда
						
						ВыполнитьОбработчик_ПВД_ПослеОбработкиПравила(ИмяПКО, ПравилоВыгрузкиДанных, ИсходящиеДанные);
						
					Иначе
						
						Выполнить(ПравилоВыгрузкиДанных.ПослеОбработки);
						
					КонецЕсли;
					
				Исключение
					ЗаписатьИнформациюОбОшибкеОбработчикиПВД(32, ПодробноеПредставлениеОшибки(ИнформацияОбОшибке()), ПравилоВыгрузкиДанных.Имя, "ПослеОбработкиВыгрузкиДанных");
				КонецПопытки;
				
			КонецЕсли;
			// {ОБРАБОТЧИК ПослеОбработки ПВД}
			
		КонецЕсли;
	
		ЗафиксироватьТранзакциюПриЗагрузкеДанных();
		
	Исключение
		
		ОтменитьТранзакциюПриЗагрузкеДанных();
		
		ВызватьИсключение(НСтр("ru = 'Ошибка при отправке данных'") + ": " + ПодробноеПредставлениеОшибки(ИнформацияОбОшибке()));
		
	КонецПопытки
	
КонецПроцедуры
//-- Проф-ИТ, #352, Корнилов М.С., 11.12.2023

#КонецОбласти