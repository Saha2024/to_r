
#Область ОбработчикиСобытийФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	
	Параметры.Свойство("ДокументРаспознан", ДокументРаспознан);
	
	Если Не Параметры.Свойство("ЭлектронныйДокумент", Документооборот)
		Или Не Параметры.Свойство("ВидДокумента", ВидДокумента) Тогда		
		Отказ = Истина;
		Возврат;
	КонецЕсли;
	
	ПодготовитьОтражениеВУчете();
	
	ОтображениеГруппыНастройкиОтраженияВУчете();
	
КонецПроцедуры

&НаКлиенте
Процедура ПриОткрытии(Отказ)
	
	УправлениеДоступностью();
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовШапкиФормы

&НаКлиенте
Процедура НадписьНастройкаОтраженияВУчетеОбработкаНавигационнойСсылки(Элемент, НавигационнаяСсылкаФорматированнойСтроки, СтандартнаяОбработка)
	СтандартнаяОбработка = Ложь;
	Элементы.ГруппаНастройкиОтраженияВУчете.Видимость = Ложь;
	
	ЗапретитьСпрашиватьПроСохранениеНастроекУчета();
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиКомандФормы

&НаКлиенте
Процедура ПерезаполнитьТекущий(Команда)
	
	Модифицированность = Ложь;
	ТекСтрока = Элементы.ЭлектронныйДокументДокументыОснования.ТекущиеДанные;
	
	Если ТекСтрока <> Неопределено Тогда
		ТекстВопроса = НСтр("ru = 'Документ учета будет заполнены данными электронного документа. Продолжить?'");
		
		СпособОбработки = ТекСтрока.СпособОбработки;
		Если ТипДокумента = ПредопределенноеЗначение("Перечисление.ТипыДокументовЭДО.УПД") Тогда
			
			СтруктураСпособовОбработки = Новый Структура("ПервичныйДокумент, СчетФактура");
			
			СпособОбработкиВторогоДокумента = "";
			Если ДокументыОснования.Количество() = 2 Тогда
				Если ТекСтрока.ОбъектУчета = ДокументыОснования[1].ОбъектУчета Тогда
					СпособОбработкиВторогоДокумента = ДокументыОснования[1].СпособОбработки;
				Иначе
					СпособОбработкиВторогоДокумента = ДокументыОснования[0].СпособОбработки;
				КонецЕсли;
			Иначе
				Если СписокТиповДокументов[0].Значение = ТекСтрока.СпособОбработки Тогда
					Если СписокТиповДокументов.Количество() > 1 Тогда
						СпособОбработкиВторогоДокумента = СписокТиповДокументов[1].Значение;
					КонецЕсли;
				Иначе
					СпособОбработкиВторогоДокумента = СписокТиповДокументов[0].Значение;
				КонецЕсли;
			КонецЕсли;
			
			Если СписокТиповДокументов[0].Значение = ТекСтрока.СпособОбработки Тогда
				СтруктураСпособовОбработки.СчетФактура = ТекСтрока.СпособОбработки;
				СтруктураСпособовОбработки.ПервичныйДокумент = СпособОбработкиВторогоДокумента;
			Иначе
				СтруктураСпособовОбработки.ПервичныйДокумент = ТекСтрока.СпособОбработки;
				СтруктураСпособовОбработки.СчетФактура = СпособОбработкиВторогоДокумента;
			КонецЕсли;
			
			СпособОбработки = СтруктураСпособовОбработки;
		КонецЕсли;
		
		СтруктураПараметров = Новый Структура("ОбъектУчета, СпособОбработки", ТекСтрока.ОбъектУчета, СпособОбработки);
		ОписаниеОповещения = Новый ОписаниеОповещения("ПерезаполнитьТекущийПродолжить", ЭтотОбъект, СтруктураПараметров);
		ПоказатьВопрос(ОписаниеОповещения, ТекстВопроса, РежимДиалогаВопрос.ДаНет, , КодВозвратаДиалога.Нет);
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ПодобратьДокумент(Команда)
	
	Модифицированность = Ложь;
	
	СпособОбработки = СтрЗаменить(Команда.Имя,"Прикрепить_","");
	НастройкиПодбора = НастройкиПодбораУчетногоДокумента(Документооборот, СпособОбработки);
	
	Если НастройкиПодбора <> Неопределено Тогда
		
		ОписаниеОповещения = Новый ОписаниеОповещения("ПодобратьДокументЗавершить", ЭтотОбъект, СпособОбработки);
		ИнтерфейсДокументовЭДОКлиент.ПоказатьПодборУчетногоДокумента(НастройкиПодбора, ОписаниеОповещения);
		
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура СоздатьДокументУчета(Команда)
	
	КонтекстОперации = ОбработкаНеисправностейБЭДКлиент.НовыйКонтекстДиагностики();
	СоздатьДокументУчетаНаСервере(СтрЗаменить(Команда.Имя,"Создать_",""), КонтекстОперации);
	ОповеститьВладельца();

	ОповещениеОЗавершении = Новый ОписаниеОповещения("СоздатьДокументУчетаЗавершение", ЭтотОбъект);
	ИнтеграцияЭДОСобытияКлиент.ПослеОтраженияВУчете(ОповещениеОЗавершении, ТипДокумента, ДокументыОснования, 
		КонтекстОперации);
	
КонецПроцедуры

&НаКлиенте
Процедура СоздатьДокументУчетаЗавершение(КонтекстДиагностики, ДополнительныеПараметры) Экспорт
	
	УправлениеДоступностью();
	ОбработкаНеисправностейБЭДКлиент.ОбработатьОшибки(КонтекстДиагностики);
	
КонецПроцедуры

&НаКлиенте
Процедура УдалитьСвязьСДокументом(Команда)
		
	ТекСтрока = Элементы.ЭлектронныйДокументДокументыОснования.ТекущиеДанные;
	
	Если ТекСтрока <> Неопределено Тогда
		ТекстВопроса = НСтр("ru = 'Связь между документами разорвется. Повторно связать документы возможно только в ручном режиме. Продолжить?'");
		ОписаниеОповещения = Новый ОписаниеОповещения("УдалитьСвязьСДокументомЗавершить", ЭтотОбъект, ТекСтрока.ОбъектУчета);
		ПоказатьВопрос(ОписаниеОповещения, ТекстВопроса, РежимДиалогаВопрос.ДаНет, , КодВозвратаДиалога.Нет);
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

&НаКлиенте
Процедура УправлениеДоступностью()

	Элементы.ФормаСоздатьДокумент.Доступность = ДокументыОснования.Количество() = 0;
	Элементы.ГруппаСоздать.Доступность = ДокументыОснования.Количество() = 0;

КонецПроцедуры

&НаСервереБезКонтекста
Функция ОбъектыУчетаДокументооборотов(Документооборот)
	
	Возврат ИнтеграцияЭДО.ОбъектыУчетаЭлектронныхДокументов(Документооборот);
		
КонецФункции

&НаКлиенте
Процедура ПодобратьДокументЗавершить(Знач Результат, Знач СпособОбработки) Экспорт
	
	Если ЗначениеЗаполнено(Результат) Тогда
		ПерепривязатьЭлектронныйДокумент(Результат,СпособОбработки);
		ОповеститьВладельца();
		УправлениеДоступностью();
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ОповеститьВладельца()
	
	Оповестить("ЭлектронныйДокументВходящий_ПодборДокументаУчета", Документооборот, ЭтотОбъект.ВладелецФормы);
	Оповестить(ИнтерфейсДокументовЭДОКлиент.ИмяСобытияОбновленияТекущихДелЭДО());
	
КонецПроцедуры

&НаСервере
Процедура СоздатьДокументУчетаНаСервере(СпособОбработки, КонтекстОперации)
	
	Если ТипДокумента = Перечисления.ТипыДокументовЭДО.УПД
		Или ТипДокумента = Перечисления.ТипыДокументовЭДО.УКД Тогда
		МассивСпособов = СтроковыеФункцииКлиентСервер.РазложитьСтрокуВМассивПодстрок(СпособОбработки, "_И_");
		Если МассивСпособов.Количество() = 2 Тогда
			СпособОбработки = Новый Структура("ПервичныйДокумент, СчетФактура", МассивСпособов[1], МассивСпособов[0]);
		КонецЕсли;
	КонецЕсли;
	
	ПараметрыЗаполнения = ЭлектронныеДокументыЭДО.ДанныеДокументовДляОтраженияВУчете(Документооборот); 
	
	Если ЗначениеЗаполнено(ПараметрыЗаполнения) Тогда
		Попытка
			НачатьТранзакцию();
			Блокировка = Новый БлокировкаДанных;
			ЭлементБлокировки = Блокировка.Добавить("Документ.ЭлектронныйДокументВходящийЭДО");
			ЭлементБлокировки.УстановитьЗначение("Ссылка", Документооборот);
			ЭлементБлокировки.Режим = РежимБлокировкиДанных.Исключительный;
			Блокировка.Заблокировать();		
			
			ОбъектыУчета = ИнтеграцияЭДО.ОбъектыУчетаАктуальногоЭлектронногоДокумента(Документооборот); 
			
			Если Не ЗначениеЗаполнено(ОбъектыУчета) Тогда 
				ИнтеграцияЭДО.ОтразитьВУчетеЭлектронныеДокументы(ПараметрыЗаполнения, СпособОбработки);
			КонецЕсли;
			ЗафиксироватьТранзакцию();
		Исключение
			ОтменитьТранзакцию();
			ВызватьИсключение;
		КонецПопытки;
		ОбновитьНастройкиОтраженияВУчете(ПараметрыЗаполнения[0], СпособОбработки);
	
	КонецЕсли;
	
	Элементы.ГруппаНастройкиОтраженияВУчете.Видимость = Ложь;
	ДокументыОснования.Загрузить(ОбъектыУчетаДокументооборотов(Документооборот));
	
КонецПроцедуры

&НаСервере
Процедура ОбновитьНастройкиОтраженияВУчете(ПараметрыЗаполнения, СпособОбработки)
	
	Если Элементы.ГруппаНастройкиОтраженияВУчете.Видимость Тогда
		
		Отказ = Ложь;
		
		КлючНастроек = НастройкиЭДОКлиентСервер.НовыйКлючНастроекОтраженияВУчете();
		ЗаполнитьЗначенияСвойств(КлючНастроек, ПараметрыЗаполнения);
	
		НастройкиЭДО.ЗаписатьНастройкуПолученияДокументов(КлючНастроек, СпособОбработки, Истина, Отказ);
		
		Если Отказ Тогда
			Возврат;
		КонецЕсли;
		
		Элементы.ГруппаНастройкиОтраженияВУчете.Видимость = Ложь;
		
	КонецЕсли;

КонецПроцедуры

&НаКлиенте
Процедура УдалитьСвязьСДокументомЗавершить(Результат, СсылкаНаОбъект) Экспорт 
	
	Если Результат = КодВозвратаДиалога.Да Тогда
		УдалитьСвязьСДокументомНаСервере(СсылкаНаОбъект, Документооборот);
		ОповеститьВладельца();
		УправлениеДоступностью();
	КонецЕсли;
	
КонецПроцедуры

&НаСервере
Процедура УдалитьСвязьСДокументомНаСервере(СсылкаНаОбъект, Документооборот)
	
	ИнтеграцияЭДО.РазорватьСвязьЭлектронногоДокументаСОбъектомУчета(Документооборот, СсылкаНаОбъект);
	ДокументыОснования.Загрузить(ОбъектыУчетаДокументооборотов(Документооборот));
	
КонецПроцедуры

&НаСервереБезКонтекста
Функция НастройкиПодбораУчетногоДокумента(Документооборот, СпособОбработки)
	
	Возврат ИнтерфейсДокументовЭДО.НастройкиПодбораУчетногоДокумента(Документооборот, СпособОбработки);
		
КонецФункции

&НаСервере
Процедура ПерепривязатьЭлектронныйДокумент(ВыбранноеЗначение, СпособОбработки)
	
	ИнтеграцияЭДО.УстановитьСвязьЭлектронногоДокументаСОбъектомУчета(Документооборот, ВыбранноеЗначение, СпособОбработки);
		
	Если Элементы.ГруппаНастройкиОтраженияВУчете.Видимость Тогда

		Отказ = Ложь;
		
		КлючНастройки = НастройкиЭДОКлиентСервер.НовыйКлючНастроекОтраженияВУчете();
		КлючНастройки.ИдентификаторОтправителя = Документооборот.ИдентификаторКонтрагента;
		КлючНастройки.ИдентификаторПолучателя = Документооборот.ИдентификаторОрганизации;
		КлючНастройки.Отправитель = Документооборот.Контрагент;
		КлючНастройки.Получатель = Документооборот.Организация;	
		КлючНастройки.ВидДокумента = Документооборот.ВидДокумента;
		
		НастройкиЭДО.ЗаписатьНастройкуПолученияДокументов(КлючНастройки, СпособОбработки, Истина, Отказ);

		Если Отказ Тогда
			Возврат;
		КонецЕсли;

		Элементы.ГруппаНастройкиОтраженияВУчете.Видимость = Ложь;

	КонецЕсли;
	
	ДокументыОснования.Загрузить(ОбъектыУчетаДокументооборотов(Документооборот));
		
КонецПроцедуры

&НаКлиенте
Процедура ЭлектронныйДокументДокументыОснованияВыбор(Элемент, ВыбраннаяСтрока, Поле, СтандартнаяОбработка)
	
	ТекСтрока = Элементы.ЭлектронныйДокументДокументыОснования.ТекущиеДанные;
	
	Если ТекСтрока <> Неопределено Тогда
		ПоказатьЗначение(,ТекСтрока.ОбъектУчета);	
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ПерезаполнитьТекущийПродолжить(Результат, СтруктураПараметров) Экспорт 
	
	Если Результат = КодВозвратаДиалога.Да Тогда
		
		КонтекстОперации = ОбработкаНеисправностейБЭДКлиент.НовыйКонтекстДиагностики();
		ПерезаполнитьОбъектУчета(СтруктураПараметров.ОбъектУчета, Документооборот,
			СтруктураПараметров.СпособОбработки);

		ОповеститьВладельца();
		
		ОповещениеОЗавершении = Новый ОписаниеОповещения("СоздатьДокументУчетаЗавершение", ЭтотОбъект);
		ИнтеграцияЭДОСобытияКлиент.ПослеПерезаполненияОбъектаУчета(ОповещениеОЗавершении,
			ТипДокумента, СтруктураПараметров.ОбъектУчета, КонтекстОперации);
		
		УправлениеДоступностью();
	КонецЕсли;
	
КонецПроцедуры

&НаСервереБезКонтекста
Процедура ПерезаполнитьОбъектУчета(ОбъектУчета, Документооборот, СпособОбработки)

	ИнтеграцияЭДО.ПерезаполнитьОбъектУчетаПоЭлектронномуДокументу(ОбъектУчета, Документооборот, СпособОбработки);
	
КонецПроцедуры

&НаСервере
Процедура ОтображениеГруппыНастройкиОтраженияВУчете()
	
	Элементы.НадписьНастройкаОтраженияВУчете.Заголовок =
		СтрШаблон(Элементы.НадписьНастройкаОтраженияВУчете.Заголовок, Документооборот.Контрагент);
	
	СпрашиватьПроСохранениеНастроекУчета = Ложь;
		
	Если ДокументРаспознан Или ЭлектронныеДокументыЭДО.ЭтоПрикладнойВидДокумента(ВидДокумента) Тогда 
	
		КлючНастройки = НастройкиЭДОКлиентСервер.НовыйКлючНастроекОтраженияВУчете();
		КлючНастройки.ИдентификаторОтправителя = Документооборот.ИдентификаторКонтрагента;
		КлючНастройки.ИдентификаторПолучателя = Документооборот.ИдентификаторОрганизации;
		КлючНастройки.Отправитель = Документооборот.Контрагент;
		КлючНастройки.Получатель = Документооборот.Организация;	
		КлючНастройки.ВидДокумента = Документооборот.ВидДокумента;
			
		НастройкиОтраженияВУчете = НастройкиЭДО.НастройкиОтраженияВУчете(КлючНастройки);
		
		СпрашиватьПроСохранениеНастроекУчета = НастройкиОтраженияВУчете.ПредлагатьСохранятьНастройки;
		
	КонецЕсли;
	
	ДокументыОснования.Загрузить(ОбъектыУчетаДокументооборотов(Документооборот));
		
	ЕстьСвязанныеДокументы = ДокументыОснования.Количество() > 0;
	Элементы.ГруппаНастройкиОтраженияВУчете.Видимость = СпрашиватьПроСохранениеНастроекУчета И Не ЕстьСвязанныеДокументы;
	
КонецПроцедуры

&НаСервере
Процедура ЗапретитьСпрашиватьПроСохранениеНастроекУчета()
				
	КлючНастройки = НастройкиЭДОКлиентСервер.НовыйКлючНастроекОтраженияВУчете();
	КлючНастройки.ИдентификаторОтправителя = Документооборот.ИдентификаторКонтрагента;
	КлючНастройки.ИдентификаторПолучателя = Документооборот.ИдентификаторОрганизации;
	КлючНастройки.Отправитель = Документооборот.Контрагент;
	КлючНастройки.Получатель = Документооборот.Организация;	
	КлючНастройки.ВидДокумента = Документооборот.ВидДокумента;
	
	Отказ = Ложь;
	
	НастройкиЭДО.ЗаписатьНастройкуПолученияДокументов(КлючНастройки, Неопределено, Истина, Отказ);
	
	Если Отказ Тогда
		Возврат;
	КонецЕсли;
	
	Элементы.ГруппаНастройкиОтраженияВУчете.Видимость = Ложь;
	
КонецПроцедуры

&НаСервере
Процедура ПодготовитьОтражениеВУчете()
	
	ДокументыОснования.Загрузить(ОбъектыУчетаДокументооборотов(Документооборот));
	
	ВидДокумента = Документооборот.ВидДокумента;
	ОписаниеВидаДокумента = ЭлектронныеДокументыЭДО.ОписаниеВидаДокумента(ВидДокумента);
	ТипДокумента = ОписаниеВидаДокумента.ТипДокумента;
	
	// Определим основной вид ЭД.
	Если ПараметрыПроизвольногоДокумента <> Неопределено
		И ПараметрыПроизвольногоДокумента.Свойство("ТипДокументаДляОтраженияВУчете") Тогда
		ТипЭлектронногоДокумента = ПараметрыПроизвольногоДокумента.ТипДокументаДляОтраженияВУчете;
		ВидДляОтраженияВУчете = ЭлектронныеДокументыЭДО.ВидДокументаПоТипу(ТипЭлектронногоДокумента);
		СписокТиповДокументов = ЭлектронныеДокументыЭДО.СписокОперацийВидаДокумента(ВидДляОтраженияВУчете);
	ИначеЕсли Не ДокументРаспознан Тогда
		СписокТиповДокументов = ИнтеграцияЭДО.СписокОперацийВсехТиповДокумента();
	Иначе
		СписокТиповДокументов = ЭлектронныеДокументыЭДО.СписокОперацийВидаДокумента(ВидДокумента);
	КонецЕсли;

	СписокТиповДокументовДоп = Новый СписокЗначений;

	Если ТипДокумента = Перечисления.ТипыДокументовЭДО.УПД Тогда
			
		ДанныеДокументооборота = ЭлектронныеДокументыЭДО.ДанныеДокументовДляОтраженияВУчете(Документооборот)[0];
		
		ДополнительныеВиды = ЭлектронныеДокументыЭДО.ДополнительныеВидыДокументовУПД(
			ДанныеДокументооборота.ДанныеОсновногоФайла);
			
		Для каждого ДополнительныйВид Из ДополнительныеВиды Цикл
			
			ДополнительныйСписок = ЭлектронныеДокументыЭДО.СписокОперацийВидаДокумента(ДополнительныйВид);
			
			Для каждого Элемент Из ДополнительныйСписок Цикл
				Если СписокТиповДокументовДоп.НайтиПоЗначению(Элемент.Значение) = Неопределено Тогда
					СписокТиповДокументовДоп.Добавить(Элемент.Значение, Элемент.Представление);
				КонецЕсли;
			КонецЦикла;	
		КонецЦикла;	
							
	ИначеЕсли ТипДокумента = Перечисления.ТипыДокументовЭДО.УКД Тогда	
		
		ДополнительныйСписок = ЭлектронныеДокументыЭДО.СписокОперацийВидаДокумента(ЭлектронныеДокументыЭДО.ВидДокументаПоТипу(
			Перечисления.ТипыДокументовЭДО.СоглашениеОбИзмененииСтоимости));
		Для каждого Элемент Из ДополнительныйСписок Цикл
			Если СписокТиповДокументовДоп.НайтиПоЗначению(Элемент.Значение) = Неопределено Тогда
				СписокТиповДокументовДоп.Добавить(Элемент.Значение, Элемент.Представление);
			КонецЕсли;
		КонецЦикла;
	КонецЕсли;
	
	Для Каждого Элемент Из СписокТиповДокументовДоп Цикл
		Элемент.Значение = СокрЛП(Элемент.Значение);
		Элемент.Представление = СокрЛП(Элемент.Представление);
	КонецЦикла;
	Для Каждого Элемент Из СписокТиповДокументов Цикл
		Элемент.Значение = СокрЛП(Элемент.Значение);
		Элемент.Представление = СокрЛП(Элемент.Представление);
	КонецЦикла;
	
	// Настроим элементы формы.
	Если СписокТиповДокументов.Количество() = 0 Тогда
		
		Элементы.ГруппаПодобрать.Видимость = Ложь;
		Элементы.ГруппаСоздать.Видимость = Ложь;
		Элементы.ФормаПодобратьДокумент.Видимость = Ложь;
		
	ИначеЕсли ТипДокумента = Перечисления.ТипыДокументовЭДО.УПД
		ИЛИ ТипДокумента = Перечисления.ТипыДокументовЭДО.УКД Тогда
		
		Элементы.ФормаПодобратьДокумент.Видимость = Ложь;
		Элементы.ФормаСоздатьДокумент.Видимость = Ложь;
		
		Для Каждого ЭлементСписка Из СписокТиповДокументов Цикл
			Для Каждого ЭлементСпискаДоп Из СписокТиповДокументовДоп Цикл
				
				НоваяКоманда = ЭтотОбъект.Команды.Добавить("Создать_" + ЭлементСписка.Значение + "_И_" + ЭлементСпискаДоп.Значение);
				НоваяКоманда.Действие = "СоздатьДокументУчета";
				
				НоваяКнопка = Элементы.Добавить(
					"Создать_" + ЭлементСписка.Значение + "_И_" + ЭлементСпискаДоп.Значение, Тип("КнопкаФормы"), Элементы.ГруппаСоздать);
				НоваяКнопка.Заголовок = ЭлементСписка.Представление + "; " + ЭлементСпискаДоп.Представление;
				НоваяКнопка.ИмяКоманды = "Создать_" + ЭлементСписка.Значение + "_И_" + ЭлементСпискаДоп.Значение;
				
			КонецЦикла;
		КонецЦикла;
		
		Для Каждого ЭлементСпискаДоп Из СписокТиповДокументовДоп Цикл
			Если СписокТиповДокументов.НайтиПоЗначению(ЭлементСпискаДоп.Значение) = Неопределено Тогда
				НовыйЭлементСписка = СписокТиповДокументов.Добавить();
				ЗаполнитьЗначенияСвойств(НовыйЭлементСписка, ЭлементСпискаДоп);
			КонецЕсли;
		КонецЦикла;
		
		Для Каждого ЭлементСписка Из СписокТиповДокументов Цикл
			НоваяКоманда = ЭтотОбъект.Команды.Добавить("Прикрепить_" + ЭлементСписка.Значение);
			НоваяКоманда.Действие = "ПодобратьДокумент";
			
			НоваяКнопка = Элементы.Добавить("Прикрепить_" + ЭлементСписка.Значение,Тип("КнопкаФормы"),Элементы.ГруппаПодобрать);
			НоваяКнопка.Заголовок = ЭлементСписка.Представление;  
			НоваяКнопка.ИмяКоманды = "Прикрепить_" + ЭлементСписка.Значение;
		КонецЦикла;
		
	ИначеЕсли СписокТиповДокументов.Количество() = 1 Тогда
		
		Элементы.ГруппаПодобрать.Видимость = Ложь;
		Элементы.ГруппаСоздать.Видимость = Ложь;
		
		НоваяКоманда = ЭтотОбъект.Команды.Добавить("Создать_" + СписокТиповДокументов[0].Значение);
		НоваяКоманда.Действие = "СоздатьДокументУчета";
		Элементы.ФормаСоздатьДокумент.ИмяКоманды = "Создать_" + СписокТиповДокументов[0].Значение;

		НоваяКоманда = ЭтотОбъект.Команды.Добавить("Прикрепить_" + СписокТиповДокументов[0].Значение);
		НоваяКоманда.Действие = "ПодобратьДокумент";
		Элементы.ФормаПодобратьДокумент.ИмяКоманды = "Прикрепить_" + СписокТиповДокументов[0].Значение;
		
	Иначе
		
		Элементы.ФормаПодобратьДокумент.Видимость = Ложь;
		Элементы.ФормаСоздатьДокумент.Видимость = Ложь;
		
		Для Каждого ЭлементСписка Из СписокТиповДокументов Цикл
			
			НоваяКоманда = ЭтотОбъект.Команды.Добавить("Создать_" + ЭлементСписка.Значение);
			НоваяКоманда.Действие = "СоздатьДокументУчета";
			
			НоваяКнопка = Элементы.Добавить("Создать_" + ЭлементСписка.Значение, Тип("КнопкаФормы"), Элементы.ГруппаСоздать);
			НоваяКнопка.Заголовок = ЭлементСписка.Представление;
			НоваяКнопка.ИмяКоманды = "Создать_" + ЭлементСписка.Значение;
			
			НоваяКоманда = ЭтотОбъект.Команды.Добавить("Прикрепить_" + ЭлементСписка.Значение);
			НоваяКоманда.Действие = "ПодобратьДокумент";
			
			НоваяКнопка = Элементы.Добавить("Прикрепить_" + ЭлементСписка.Значение, Тип("КнопкаФормы"), Элементы.ГруппаПодобрать);
			НоваяКнопка.Заголовок = ЭлементСписка.Представление;  
			НоваяКнопка.ИмяКоманды = "Прикрепить_" + ЭлементСписка.Значение;
			
		КонецЦикла;
		
	КонецЕсли;
	
	Если ЭлектронныеДокументыЭДО.ЭтоПрикладнойТипДокумента(ВидДокумента) Тогда
		АктуальныеВидыЭД = ЭлектронныеДокументыЭДО.ПрикладныеВидыДокументов();
	Иначе
		АктуальныеВидыЭД = ЭлектронныеДокументыЭДО.ИспользуемыеВидыДокументовВходящие();
	КонецЕсли;
	
	Если АктуальныеВидыЭД.Найти(ВидДокумента) = Неопределено Или Не ДокументРаспознан Тогда
		Элементы.ФормаСоздатьДокумент.Видимость = Ложь;
		Элементы.ГруппаСоздать.Видимость = Ложь;
		Элементы.ФормаПерезаполнитьТекущий.Видимость = Ложь;
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти


