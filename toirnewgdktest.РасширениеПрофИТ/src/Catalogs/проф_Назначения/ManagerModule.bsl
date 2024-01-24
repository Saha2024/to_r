//++ Проф-ИТ, #359, Соловьев А.А., 23.11.2023

#Область СозданиеНазначений

Функция ПроверитьЗаполнитьПередЗаписью(Назначение, ШаблонНазначения, ЭтотОбъект, РеквизитыПроверки, Отказ, ДляНовыхСоздаватьВсегда = Ложь, Очистить = Ложь) Экспорт
	
	// Проверяет, что назначение в объекте корректное.
	// Если назначение некорректное, то исправляет его:
	//  - либо ищет корректное назначение в базе,
	//  - либо выдает ссылку нового назначения, если поиск не даст результатов.
	
	ОбновленоНазначение = Ложь;
	
	// Для нового объекта всегда выполняется обновление назначения.
	Если ЭтотОбъект.ЭтоНовый() Тогда
		
		ОбновленоНазначение = Истина;
		
		// Назначение в объекте не используется, значит нужно его очистить.
		Если Очистить Тогда
			
			Назначение = Неопределено;
			
		// В шаблоне есть ссылка на создаваемый объект, значит искать корректное назначение в базе нет смысла, его там нет.
		// Назначение нужно будет создать при записи объекта.
		ИначеЕсли ДляНовыхСоздаватьВсегда Или СсылкаЕстьВШаблоне(ЭтотОбъект.Ссылка, ШаблонНазначения) Тогда
			
			Назначение = Справочники.проф_Назначения.ПолучитьСсылку();
			
		// Назначением служит назначение, связанное с другим объектом, значит оно есть в базе. Нужно найти его.
		Иначе
			
			// Назначение связанное с направлением деятельности создаются только в момент фиксации потребности.
			Если Не (ШаблонНазначения.НаправлениеДеятельности.Пустая()
					И ШаблонНазначения.ТипНазначения = Перечисления.проф_ТипыНазначений.Собственное) Тогда
				
				// Установим блокировку для исключения возможности создания дубля другим объектом, генерирующим в это время такое же назначение.
				УстановитьБлокировкуПоШаблону(ШаблонНазначения);
				
				Назначение = НайтиПоШаблону(ШаблонНазначения);
				
				Если Не ЗначениеЗаполнено(Назначение) Тогда
					
					Назначение = Справочники.проф_Назначения.ПолучитьСсылку();
					
				КонецЕсли;
				
			Иначе
				
				Назначение = НайтиПоШаблону(ШаблонНазначения);
				
				Если Не ЗначениеЗаполнено(Назначение) Тогда
					
					Отказ = Истина;
					ОбщегоНазначенияКлиентСервер.СообщитьПользователю(СообщениеНетНазначенияВБазе(ШаблонНазначения));
					
				КонецЕсли;
				
			КонецЕсли;
			
		КонецЕсли;
		
	// Если объект записывается повторно, назначение следует обновлять, если изменились ключевые реквизиты объекта.
	ИначеЕсли РеквизитыПроверки <> "" Тогда
		
		ОбъектВБазе = ОбщегоНазначения.ЗначенияРеквизитовОбъекта(ЭтотОбъект.Ссылка, РеквизитыПроверки);
		Если Не ОбщегоНазначенияУТКлиентСервер.СтруктурыРавны(ЭтотОбъект, ОбъектВБазе, РеквизитыПроверки)
				Или Не ЗначениеЗаполнено(Назначение) Тогда // назначение могло очиститься если ключевые реквизиты меняли в форме после переключения опции.
			
			ОбновленоНазначение = Истина;
			
			// Назначение в объекте не используется, значит нужно его очистить.
			Если Очистить Тогда
				
				Назначение = Неопределено;
				
			// Назначение уже могло быть записано в базу если реквизиты объекта меняют "туда-обратно" и такая комбинация уже была в объекте.
			// Если же нет, то его нужно будет создать при записи объекта.
			ИначеЕсли СсылкаЕстьВШаблоне(ЭтотОбъект.Ссылка, ШаблонНазначения) Тогда
				
				Назначение = НайтиПоШаблону(ШаблонНазначения);
				Если Не ЗначениеЗаполнено(Назначение) Тогда
					Назначение = Справочники.проф_Назначения.ПолучитьСсылку(); // это запись объекта с новыми значениями ключевых реквизитов
				КонецЕсли;
				
			// Назначением служит назначение, связанное с другим объектом, значит оно есть в базе. Нужно найти его.
			Иначе
				
				// Назначение связанное с направлением деятельности создаются только в момент фиксации потребности.
				Если Не (ШаблонНазначения.НаправлениеДеятельности.Пустая()
						И ШаблонНазначения.ТипНазначения = Перечисления.проф_ТипыНазначений.Собственное) Тогда
						
					// Установим блокировку для исключения возможности создания дубля другим объектом, генерирующим в это время такое же назначение.
					УстановитьБлокировкуПоШаблону(ШаблонНазначения);
					
					Назначение = НайтиПоШаблону(ШаблонНазначения);
					
					Если Не ЗначениеЗаполнено(Назначение) Тогда
						Назначение = Справочники.проф_Назначения.ПолучитьСсылку();
					КонецЕсли;
					
				Иначе
					
					Назначение = НайтиПоШаблону(ШаблонНазначения);
					
					Если Не ЗначениеЗаполнено(Назначение) Тогда
						
						Отказ = Истина;
						ОбщегоНазначенияКлиентСервер.СообщитьПользователю(СообщениеНетНазначенияВБазе(ШаблонНазначения));
						
					КонецЕсли;
					
				КонецЕсли;
				
			КонецЕсли;
			
		КонецЕсли;
		
	КонецЕсли;
	
	Возврат ОбновленоНазначение;
	
КонецФункции

Процедура ПриЗаписиДокумента(Назначение, ШаблонНазначения, ЭтотОбъект, Получатель, ВидДеятельностиНДС) Экспорт
	
	Если Не ЗначениеЗаполнено(Назначение) Тогда
		Возврат;
	КонецЕсли;
	
	РеквизитыДокумента = Новый Структура("Номер, Дата, ПометкаУдаления, Подразделение");
	ЗаполнитьЗначенияСвойств(РеквизитыДокумента, ЭтотОбъект);
	РеквизитыДокумента.Вставить("Получатель",         Получатель);
	РеквизитыДокумента.Вставить("ВидДеятельностиНДС", ВидДеятельностиНДС);
	
	УстановитьПривилегированныйРежим(Истина);
	РеквизитыНазначения = ОбщегоНазначения.ЗначенияРеквизитовОбъекта(Назначение,
		"Наименование, ПометкаУдаления, Заказ, НаправлениеДеятельности, ВидДеятельностиНДС, ТипНазначения, ТипОбъектаНазначения");
	
	Если РеквизитыНазначения.Наименование <> Неопределено Тогда
		
		ТипОбъектаНазначения = РеквизитыНазначения.ТипОбъектаНазначения;
		Если Не ЗначениеЗаполнено(РеквизитыНазначения.ТипОбъектаНазначения) Тогда
			ОпределитьТипОбъектаНазначения(РеквизитыНазначения);
		КонецЕсли;
		
		ПредставлениеНазначения = ПредставлениеНазначенияДляЗаписиВИнформационнуюБазу(
			РеквизитыНазначения,
			РеквизитыДокумента);
		
		ОбновитьПометкуУдаления    = ЗначениеЗаполнено(РеквизитыНазначения.Заказ)
			И РеквизитыНазначения.ПометкаУдаления <> РеквизитыДокумента.ПометкаУдаления;
		ОбновитьВидДеятельностиНДС = ЗначениеЗаполнено(РеквизитыНазначения.Заказ)
			И РеквизитыНазначения.ВидДеятельностиНДС <> РеквизитыДокумента.ВидДеятельностиНДС;
		ОбновитьНаименование = РеквизитыНазначения.Наименование <> ПредставлениеНазначения;
		ОбновитьТипОбъектаНазначения = РеквизитыНазначения.ТипОбъектаНазначения <> ТипОбъектаНазначения;

		Если ОбновитьПометкуУдаления
			Или ОбновитьВидДеятельностиНДС
			Или ОбновитьНаименование
			Или ОбновитьТипОбъектаНазначения Тогда
			
			НазначениеОбъект = Назначение.ПолучитьОбъект();
			
			Если ОбновитьПометкуУдаления Тогда
				НазначениеОбъект.ПометкаУдаления = РеквизитыДокумента.ПометкаУдаления;
			КонецЕсли;
			
			Если ОбновитьВидДеятельностиНДС Тогда
				НазначениеОбъект.ВидДеятельностиНДС = РеквизитыДокумента.ВидДеятельностиНДС;
			КонецЕсли;
			
			Если ОбновитьНаименование Тогда
				НазначениеОбъект.Наименование = ПредставлениеНазначения;
			КонецЕсли;
			
			Если ОбновитьТипОбъектаНазначения Тогда
				НазначениеОбъект.ТипОбъектаНазначения = РеквизитыНазначения.ТипОбъектаНазначения;
			КонецЕсли;

			НазначениеОбъект.Записать();
			
		КонецЕсли;
		
	Иначе // Это новое назначение
		
		ШаблонНазначения.ТипОбъектаНазначения = ТипОбъектаНазначенияПоДокументу(ШаблонНазначения.Заказ);
		
		ПредставлениеНазначения = ПредставлениеНазначенияДляЗаписиВИнформационнуюБазу(
			ШаблонНазначения,
			РеквизитыДокумента);
		
		НазначениеОбъект = Справочники.проф_Назначения.СоздатьЭлемент();
		НазначениеОбъект.УстановитьСсылкуНового(Назначение);
		
		НазначениеОбъект.НаправлениеДеятельности	= ШаблонНазначения.НаправлениеДеятельности;
		НазначениеОбъект.Заказ						= ШаблонНазначения.Заказ;
		НазначениеОбъект.ТипНазначения				= ШаблонНазначения.ТипНазначения;
		НазначениеОбъект.Наименование				= ПредставлениеНазначения;
		Если ЗначениеЗаполнено(ШаблонНазначения.Заказ) Тогда
			НазначениеОбъект.ПометкаУдаления		= РеквизитыДокумента.ПометкаУдаления;
			НазначениеОбъект.ВидДеятельностиНДС		= РеквизитыДокумента.ВидДеятельностиНДС;
			НазначениеОбъект.Подразделение			= РеквизитыДокумента.Подразделение;
			НазначениеОбъект.ИДЗаказа				= ШаблонНазначения.Заказ.УникальныйИдентификатор();
		КонецЕсли;
		НазначениеОбъект.ТипОбъектаНазначения = ШаблонНазначения.ТипОбъектаНазначения;
		НазначениеОбъект.ДвиженияПоСкладскимРегистрам = Истина;
		НазначениеОбъект.Записать();
		
	КонецЕсли;
	
КонецПроцедуры

Процедура ПриЗаписиСправочника(Назначение, ШаблонНазначения, ЭтотОбъект, ВидДеятельностиНДС, ОбновитьНазначение = Ложь) Экспорт
	
	Если Не ЗначениеЗаполнено(Назначение) Тогда
		Возврат;
	КонецЕсли;
	
	УстановитьПривилегированныйРежим(Истина);
	РеквизитыНазначения = ОбщегоНазначения.ЗначенияРеквизитовОбъекта(Назначение,
		"Наименование, ПометкаУдаления, Заказ, НаправлениеДеятельности, ВидДеятельностиНДС, ТипНазначения, ТипОбъектаНазначения");
	
	Если РеквизитыНазначения.Наименование <> Неопределено Тогда
		
		ТипОбъектаНазначения = РеквизитыНазначения.ТипОбъектаНазначения;
		Если Не ЗначениеЗаполнено(РеквизитыНазначения.ТипОбъектаНазначения) Тогда
			ОпределитьТипОбъектаНазначения(РеквизитыНазначения);
		КонецЕсли;
		
		ПредставлениеНазначения = ПредставлениеНазначенияДляЗаписиВИнформационнуюБазу(
			РеквизитыНазначения,
			Неопределено);
		
		ОбновитьПометкуУдаления    = РеквизитыНазначения.ПометкаУдаления <> ЭтотОбъект.ПометкаУдаления;
		ОбновитьВидДеятельностиНДС = РеквизитыНазначения.ВидДеятельностиНДС <> ВидДеятельностиНДС;
		ОбновитьНаименование       = РеквизитыНазначения.Наименование <> ПредставлениеНазначения;
		ОбновитьТипОбъектаНазначения = РеквизитыНазначения.ТипОбъектаНазначения <> ТипОбъектаНазначения;
		
		Если ОбновитьНазначение
			Или ОбновитьПометкуУдаления
			Или ОбновитьВидДеятельностиНДС
			Или ОбновитьНаименование
			Или ОбновитьТипОбъектаНазначения Тогда
			
			НазначениеОбъект = Назначение.ПолучитьОбъект();
			
			Если ОбновитьПометкуУдаления Тогда
				НазначениеОбъект.ПометкаУдаления = ЭтотОбъект.ПометкаУдаления;
			КонецЕсли;
			
			Если ОбновитьВидДеятельностиНДС Тогда
				НазначениеОбъект.ВидДеятельностиНДС = ВидДеятельностиНДС;
			КонецЕсли;
			
			Если ОбновитьНаименование Тогда
				НазначениеОбъект.Наименование = ПредставлениеНазначения;
			КонецЕсли;
			
			Если ОбновитьТипОбъектаНазначения Тогда
				НазначениеОбъект.ТипОбъектаНазначения = РеквизитыНазначения.ТипОбъектаНазначения;
			КонецЕсли;
			
			НазначениеОбъект.Записать();
			
		КонецЕсли;
		
	Иначе // Это новое назначение
		
		ОпределитьТипОбъектаНазначения(ШаблонНазначения);
		
		ПредставлениеНазначения = ПредставлениеНазначенияДляЗаписиВИнформационнуюБазу(
			ШаблонНазначения,
			Неопределено);
		
		НазначениеОбъект = Справочники.проф_Назначения.СоздатьЭлемент();
		НазначениеОбъект.УстановитьСсылкуНового(Назначение);
		
		НазначениеОбъект.НаправлениеДеятельности	= ШаблонНазначения.НаправлениеДеятельности;
		НазначениеОбъект.Заказ						= ШаблонНазначения.Заказ;
		НазначениеОбъект.ТипНазначения				= ШаблонНазначения.ТипНазначения;
		НазначениеОбъект.Наименование				= ПредставлениеНазначения;
		НазначениеОбъект.ПометкаУдаления			= ЭтотОбъект.ПометкаУдаления;
		НазначениеОбъект.ВидДеятельностиНДС			= ВидДеятельностиНДС;
		НазначениеОбъект.ТипОбъектаНазначения		= ШаблонНазначения.ТипОбъектаНазначения;
		
		НазначениеОбъект.Записать();
		
	КонецЕсли;
	
КонецПроцедуры

Функция ПредставлениеНазначенияДляЗаписиВИнформационнуюБазу(КлючевыеПоля, РеквизитыДокумента) Экспорт
	
	ПредставлениеНазначения = Неопределено;
	
	МассивПолей = Новый Массив;
	
	Если ЗначениеЗаполнено(КлючевыеПоля.Заказ) Тогда
		
		МассивПолей.Добавить(РеквизитыДокумента.Получатель);
		
		Если ЗначениеЗаполнено(КлючевыеПоля.НаправлениеДеятельности) Тогда
			МассивПолей.Добавить(КлючевыеПоля.НаправлениеДеятельности);
		КонецЕсли;
		
		НомерЗаказа = ПрефиксацияОбъектовКлиентСервер.УдалитьЛидирующиеНулиИзНомераОбъекта(РеквизитыДокумента.Номер);
		МассивПолей.Добавить(НомерЗаказа);
		
		МассивПолей.Добавить(Формат(РеквизитыДокумента.Дата, "ДЛФ=D"));
		
	Иначе
		
		МассивПолей.Добавить(КлючевыеПоля.НаправлениеДеятельности);
		
	КонецЕсли;
	
	ПредставлениеНазначения = СтрСоединить(МассивПолей, ", ");
	
	ТипОбъектаНазначенияНаОсновномЯзыке = ТипОбъектаНазначенияНаОсновномЯзыке(КлючевыеПоля.ТипОбъектаНазначения);
	ПредставлениеНазначения = ПредставлениеНазначения + " (" + ТипОбъектаНазначенияНаОсновномЯзыке + ")";
	
	Возврат ПредставлениеНазначения;
	
КонецФункции

Функция ШаблоныРавны(Шаблон1, Шаблон2) Экспорт
	
	Для Каждого Свойство Из Шаблон1 Цикл
		
		Если Шаблон1[Свойство.Ключ] <> Шаблон2[Свойство.Ключ] Тогда
			Возврат Ложь;
		КонецЕсли;
		
	КонецЦикла;
	Возврат Истина;
	
КонецФункции

// Возвращаемое значение:
//  Структура - структура:
//   * НаправлениеДеятельности - СправочникСсылка.проф_НаправленияДеятельности -
//   * Заказ - ОпределяемыйТип.ОжидаемаяОтгрузкаИлиПоступление, Неопределено -
//   * ТипНазначения - ПеречислениеСсылка.ТипыНазначений -
//   * ТипОбъектаНазначения - ПеречислениеСсылка.ТипыОбъектовНазначений -
Функция ШаблонНового() Экспорт
	
	Результат = Новый Структура();
	Результат.Вставить("НаправлениеДеятельности", Справочники.проф_НаправленияДеятельности.ПустаяСсылка());
	Результат.Вставить("Заказ", Неопределено);
	Результат.Вставить("ТипНазначения", Перечисления.проф_ТипыНазначений.Собственное);
	Результат.Вставить("ТипОбъектаНазначения", Перечисления.проф_ТипОбъектаНазначения.ПустаяСсылка());
	
	Возврат Результат;
	
КонецФункции

Функция НайтиПоШаблону(Шаблон) Экспорт
	
	Результат = Неопределено;
	
	Запрос = Новый Запрос();
	Запрос.Текст =
	"ВЫБРАТЬ ПЕРВЫЕ 1
	|	Назначения.Ссылка КАК Ссылка
	|ИЗ
	|	Справочник.проф_Назначения КАК Назначения
	|ГДЕ
	|	Назначения.НаправлениеДеятельности = &НаправлениеДеятельности
	|	И Назначения.Заказ = &Заказ
	|	И Назначения.ТипНазначения = &ТипНазначения";
	
	Запрос.УстановитьПараметр("НаправлениеДеятельности", Шаблон.НаправлениеДеятельности);
	Запрос.УстановитьПараметр("Заказ",                   Шаблон.Заказ);
	Запрос.УстановитьПараметр("ТипНазначения",           Шаблон.ТипНазначения);
	
	УстановитьПривилегированныйРежим(Истина);
	Выборка = Запрос.Выполнить().Выбрать();
	Если Выборка.Следующий() Тогда
		Результат = Выборка.Ссылка;
	КонецЕсли;
	
	Возврат Результат;
	
КонецФункции

Процедура УстановитьБлокировкуПоШаблону(Шаблон) Экспорт
	
	Блокировка = Новый БлокировкаДанных;
	ЭлементБлокировки = Блокировка.Добавить("Справочник.проф_Назначения");
	Для каждого ПолеШаблона Из Шаблон Цикл
		ЭлементБлокировки.УстановитьЗначение(ПолеШаблона.Ключ, ПолеШаблона.Значение);
	КонецЦикла;
	Блокировка.Заблокировать();
	
КонецПроцедуры

Процедура ОпределитьТипОбъектаНазначения(РеквизитыНазначения)
	
	Если ЗначениеЗаполнено(РеквизитыНазначения.Заказ) Тогда
		
		РеквизитыНазначения.ТипОбъектаНазначения = ТипОбъектаНазначенияПоДокументу(РеквизитыНазначения.Заказ);
		
	Иначе
		
		РеквизитыНазначения.ТипОбъектаНазначения = Перечисления.проф_ТипОбъектаНазначения.НаправлениеДеятельности;
		
	КонецЕсли;

КонецПроцедуры

Функция ТипОбъектаНазначенияПоДокументу(Документ)
	
	ЗначениеПеречисления = Перечисления.проф_ТипОбъектаНазначения.ПустаяСсылка();
	
	Если ТипЗнч(Документ) = Тип("ДокументСсылка.торо_ЗаявкаНаРемонт") Тогда
		ЗначениеПеречисления = Перечисления.проф_ТипОбъектаНазначения.ЗаявкаНаРемонт;
	ИначеЕсли ТипЗнч(Документ) = Тип("ДокументСсылка.ЗаказНаВнутреннееПотребление") Тогда
		ЗначениеПеречисления = Перечисления.проф_ТипОбъектаНазначения.ЗаказНаВнутреннееПотребление;
	КонецЕсли;
	
	Возврат ЗначениеПеречисления;
	
КонецФункции

Функция СообщениеНетНазначенияВБазе(ШаблонНазначения)
	
	Сообщение = Неопределено;
	Если ЗначениеЗаполнено(ШаблонНазначения.НаправлениеДеятельности) Тогда
		
		Сообщение = СтрШаблон(НСтр("ru = 'В информационной базе не обнаружено назначение для направления деятельности: ""%1""';
									|en = 'Assignment for line of business ""%1"" was not found in the infobase'"),
			ШаблонНазначения.НаправлениеДеятельности);
			
	Иначе
		
		Сообщение = СтрШаблон(НСтр("ru = 'В информационной базе не обнаружено нужное назначение';
									|en = 'Required assignment is not found in the infobase'"));
		
	КонецЕсли;
	
	Возврат Сообщение;
	
КонецФункции

Функция СсылкаЕстьВШаблоне(Ссылка, Шаблон)
	
	Для Каждого Свойство Из Шаблон Цикл
		
		Если Свойство.Значение = Ссылка Тогда
			Возврат Истина;
		КонецЕсли;
		
	КонецЦикла;
	
	Возврат Ложь;
	
КонецФункции

Функция ТипОбъектаНазначенияНаОсновномЯзыке(ТипОбъектаНазначения)
	
	КодОсновногоЯзыка = ОбщегоНазначения.КодОсновногоЯзыка();
	
	ПредставлениеТипаОбъектаНазначения = "";
	
	Если ТипОбъектаНазначения = Перечисления.проф_ТипОбъектаНазначения.ЗаявкаНаРемонт Тогда
		ПредставлениеТипаОбъектаНазначения = НСтр("ru = 'Сметы ремонта (Заявки на ремонт)'", КодОсновногоЯзыка);
	ИначеЕсли ТипОбъектаНазначения = Перечисления.проф_ТипОбъектаНазначения.ЗаказНаВнутреннееПотребление Тогда
		ПредставлениеТипаОбъектаНазначения = НСтр("ru = 'Заказ на внутреннее потребление'", КодОсновногоЯзыка);
	КонецЕсли;
	
	Возврат ПредставлениеТипаОбъектаНазначения;
	
КонецФункции

#КонецОбласти

//-- Проф-ИТ, #359, Соловьев А.А., 23.11.2023