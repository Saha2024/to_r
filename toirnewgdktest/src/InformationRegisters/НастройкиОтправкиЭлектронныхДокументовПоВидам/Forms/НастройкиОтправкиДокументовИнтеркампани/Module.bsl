
#Область ОбработчикиСобытийФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	
	ТолькоПросмотр =  Не НастройкиЭДО.ЕстьПравоНастройкиОбмена();
	
	Параметры.Свойство("Отправитель", Отправитель);
	Параметры.Свойство("Получатель" , Получатель);
	
	Элементы.ГруппаОжиданиеУдаления.Видимость = Ложь;
	
	Создание = Неопределено;
	Параметры.Свойство("Создание" , Создание);
	
	Если Создание = Истина Тогда
		Элементы.Получатель.ТолькоПросмотр = Ложь;
		Элементы.Отправитель.ТолькоПросмотр = ЗначениеЗаполнено(Отправитель);
		Заголовок = СтрШаблон(НСтр("ru = '%1 (Создание)'"), Заголовок);
	КонецЕсли;
	
	ТаблицаДокументов = СформироватьТаблицуНастроек(ИсходящиеДокументы, Отправитель, Получатель);
	
	ИсходящиеДокументы.Загрузить(ТаблицаДокументов);
	
	// ЭлектронноеВзаимодействие.ОбменСКонтрагентами.КонтекстныеПодсказкиБЭД
	КонтекстныеПодсказкиБЭД.КонтекстныеПодсказки_ПриСозданииНаСервере(ЭтотОбъект, 
																		Элементы.ПанельКонтекстныхНовостей, 
																		Элементы.ГруппаКонтекстныхПодсказок);
	СформироватьКонтекст();
	// Конец ЭлектронноеВзаимодействие.ОбменСКонтрагентами.КонтекстныеПодсказкиБЭД

КонецПроцедуры

&НаКлиенте
Процедура ПриОткрытии(Отказ)
	
	// ЭлектронноеВзаимодействие.ОбменСКонтрагентами.КонтекстныеПодсказкиБЭД
	КонтекстныеПодсказкиБЭДКлиент.КонтекстныеНовости_ПриОткрытии(ЭтотОбъект);
	// Конец ЭлектронноеВзаимодействие.ОбменСКонтрагентами.КонтекстныеПодсказкиБЭД

КонецПроцедуры

&НаКлиенте
Процедура ПередЗакрытием(Отказ, ЗавершениеРаботы, ТекстПредупреждения, СтандартнаяОбработка)
	
	Если Модифицированность Тогда
		
		Отказ = Истина;
		
		Если ЗавершениеРаботы Тогда
			ТекстПредупреждения = НСтр("ru = 'Настройки отправки не сохранены. Завершить работу с программой?'")
		Иначе
			Описание = Новый ОписаниеОповещения("ПередЗакрытиемОкнаФормы", ЭтотОбъект);
			ПоказатьВопрос(Описание, НСтр("ru = 'Настройки отправки изменены. Сохранить изменения?'"), РежимДиалогаВопрос.ДаНетОтмена);
		КонецЕсли;
		
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ОбработкаОповещения(ИмяСобытия, Параметр, Источник)
	
	// ЭлектронноеВзаимодействие.ОбменСКонтрагентами.КонтекстныеПодсказкиБЭД
	КонтекстныеПодсказкиБЭДКлиент.КонтекстныеНовости_ОбработкаОповещения(ЭтотОбъект, ИмяСобытия, Параметр, Источник);
	// Конец ЭлектронноеВзаимодействие.ОбменСКонтрагентами.КонтекстныеПодсказкиБЭД

КонецПроцедуры

#КонецОбласти

#Область ОбработчикиКомандФормы

&НаКлиенте
Процедура Записать(Команда)
	
	Модифицированность = НЕ НачатьСохраненияНастроек();
	
КонецПроцедуры

&НаКлиенте
Процедура ЗаписатьИЗакрыть(Команда)
	
	Если НачатьСохраненияНастроек() Тогда
		Модифицированность = Ложь;
		Закрыть();
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура Удалить(Команда)
	
	Описание = Новый ОписаниеОповещения("УдалитьНастройкиОтправкиИнтеркампани", ЭтотОбъект);
	
	ТекстВопроса = НСтр("ru = 'Сейчас будет удалена настройка обмена интеркампани.
                         |Продолжить?'");
	
	ПоказатьВопрос(Описание, ТекстВопроса, РежимДиалогаВопрос.ДаНет);

КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

&НаСервереБезКонтекста
Функция СформироватьТаблицуНастроек(Знач ИсходящиеДокументы, Знач Отправитель, Знач Получатель)
	
	ИсходящиеДокументы = ИсходящиеДокументы.Выгрузить();
	ИсходящиеДокументы.Очистить();
	
	ВидыДокументов = ЭлектронныеДокументыЭДО.ИспользуемыеВидыДокументовИнтеркампани();
	
	НастройкиОтправкиЭДОСлужебный.ЗаполнитьНастройкиПоВидамЭлектронныхДокументовИнтеркампани(ИсходящиеДокументы,
		ВидыДокументов, Отправитель, Получатель);
	
	Запрос = Новый Запрос;
	Запрос.Текст = 
		"ВЫБРАТЬ
		|	НастройкиЭДО.Отправитель КАК Отправитель,
		|	НастройкиЭДО.Получатель КАК Получатель,
		|	НастройкиЭДО.ВидДокумента КАК ВидДокумента,
		|	НастройкиЭДО.ИдентификаторОтправителя КАК ИдентификаторОтправителя,
		|	НастройкиЭДО.ИдентификаторПолучателя КАК ИдентификаторПолучателя,
		|	НастройкиЭДО.СпособОбменаЭД КАК СпособОбменаЭД,
		|	НастройкиЭДО.МаршрутПодписания КАК МаршрутПодписания,
		|	НастройкиЭДО.Формировать КАК Формировать,
		|	НастройкиЭДО.ВерсияФормата КАК ВерсияФормата
		|ПОМЕСТИТЬ НастройкиЭДО
		|ИЗ
		|	&НастройкиЭДО КАК НастройкиЭДО
		|;
		|
		|////////////////////////////////////////////////////////////////////////////////
		|ВЫБРАТЬ
		|	НастройкиЭДО.Отправитель КАК Отправитель,
		|	НастройкиЭДО.Получатель КАК Получатель,
		|	НастройкиЭДО.ВидДокумента КАК ВидДокумента,
		|	ЕСТЬNULL(НастройкиОтправкиЭлектронныхДокументовПоВидам.ИдентификаторОтправителя, НастройкиЭДО.ИдентификаторОтправителя) КАК ИдентификаторОтправителя,
		|	ЕСТЬNULL(НастройкиОтправкиЭлектронныхДокументовПоВидам.ИдентификаторПолучателя, НастройкиЭДО.ИдентификаторПолучателя) КАК ИдентификаторПолучателя,
		|	НастройкиЭДО.СпособОбменаЭД КАК СпособОбменаЭД,
		|	ЕСТЬNULL(НастройкиОтправкиЭлектронныхДокументовПоВидам.МаршрутПодписания, НастройкиЭДО.МаршрутПодписания) КАК МаршрутПодписания,
		|	ЕСТЬNULL(НастройкиОтправкиЭлектронныхДокументовПоВидам.Формировать, НастройкиЭДО.Формировать) КАК Формировать,
		|	ЕСТЬNULL(НастройкиОтправкиЭлектронныхДокументовПоВидам.ВерсияФормата, НастройкиЭДО.ВерсияФормата) КАК ВерсияФормата
		|ИЗ
		|	НастройкиЭДО КАК НастройкиЭДО
		|		ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.НастройкиОтправкиЭлектронныхДокументовПоВидам КАК НастройкиОтправкиЭлектронныхДокументовПоВидам
		|		ПО НастройкиЭДО.Отправитель = НастройкиОтправкиЭлектронныхДокументовПоВидам.Отправитель
		|			И НастройкиЭДО.Получатель = НастройкиОтправкиЭлектронныхДокументовПоВидам.Получатель
		|			И НастройкиЭДО.ВидДокумента = НастройкиОтправкиЭлектронныхДокументовПоВидам.ВидДокумента";
	
	Запрос.УстановитьПараметр("НастройкиЭДО", ИсходящиеДокументы);
	
	Возврат Запрос.Выполнить().Выгрузить();
	
КонецФункции

&НаКлиенте
Функция НачатьСохраненияНастроек()
	
	Если Не ПроверитьЗаполнение() Тогда
		Возврат Ложь;
	КонецЕсли;
	
	Результат = ЗаписатьНастройки(Отправитель, Получатель, ИсходящиеДокументы);
	
	Если Результат Тогда
		Оповестить("ОбновитьСостояниеЭД");
		ТекстЗаголовка = НСтр("ru = 'Успех'");
		ТекстСообщения = НСтр("ru = 'Настройки обмена между нашими организациями сохранены успешно'");
		ПоказатьОповещениеПользователя(ТекстЗаголовка,, ТекстСообщения);
	КонецЕсли;
	
	Возврат Результат;

КонецФункции

&НаСервереБезКонтекста
Функция ЗаписатьНастройки(Знач Отправитель, Знач Получатель, Знач ИсходящиеДокументы)
	
	ИсходящиеДокументы = ИсходящиеДокументы.Выгрузить();
	
	Возврат НастройкиОтправкиЭДОСлужебный.ЗаписатьНастройкиИнтеркампани(ИсходящиеДокументы, Отправитель, Получатель);
	
КонецФункции

&НаКлиенте
Процедура УдалитьНастройкиОтправкиИнтеркампани(РезультатВопроса, ДополнительныеПараметры) Экспорт
	
	Если РезультатВопроса <> КодВозвратаДиалога.Да Тогда
		Возврат;
	КонецЕсли;
	
	Модифицированность = Ложь;
	
	ПараметрыПроцедуры = Новый Структура;
	ПараметрыПроцедуры.Вставить("Отправитель"       , Отправитель);
	ПараметрыПроцедуры.Вставить("Получатель"        , Получатель);
	
	ДлительнаяОперация = НачатьУдалениеНастроекОтправкиНаСервере(УникальныйИдентификатор, ПараметрыПроцедуры);
	
	Если ДлительнаяОперация <> Неопределено Тогда
		
		ТолькоПросмотр = Истина;
		Элементы.ГруппаОжиданиеУдаления.Видимость = Истина;
		
		ПараметрыОжидания = ДлительныеОперацииКлиент.ПараметрыОжидания(ЭтаФорма);
		ПараметрыОжидания.ВыводитьОкноОжидания = Ложь;
		
		Описание = Новый ОписаниеОповещения("ПриЗавершенииУдаленияНастроекОтправки", ЭтаФорма);
		ДлительныеОперацииКлиент.ОжидатьЗавершение(ДлительнаяОперация, Описание, ПараметрыОжидания);
		
	КонецЕсли;
	
КонецПроцедуры

&НаСервереБезКонтекста
Функция НачатьУдалениеНастроекОтправкиНаСервере(Знач УникальныйИдентификатор, Знач ПараметрыПроцедуры)
	
	ПараметрыВыполнения = ДлительныеОперации.ПараметрыВыполненияВФоне(УникальныйИдентификатор);
	ПараметрыВыполнения.НаименованиеФоновогоЗадания = НСтр("ru = 'Удаление настройки интеркампани'");
	
	Возврат ДлительныеОперации.ВыполнитьВФоне("РегистрыСведений.НастройкиОтправкиЭлектронныхДокументовПоВидам.УдалитьНастройкиИнтеркампани",
		ПараметрыПроцедуры, ПараметрыВыполнения);
	
КонецФункции

&НаКлиенте
Процедура ПриЗавершенииУдаленияНастроекОтправки(Результат, ДополнительныеПараметры) Экспорт
	
	ТолькоПросмотр = Ложь;
	Элементы.ГруппаОжиданиеУдаления.Видимость = Ложь;

	Если Результат = Неопределено Тогда
		Возврат;
	КонецЕсли;
	
	Если Результат.Статус = "Ошибка" Тогда
		
		ТекстСообщения = НСтр("ru = 'Во время удаления настройки интеркампани произошла ошибка.'");
		
		ОбработкаНеисправностейБЭДВызовСервера.ОбработатьОшибку(НСтр("ru = 'Удаление настройки интеркампани'"), 
			ОбщегоНазначенияБЭДКлиентСервер.ПодсистемыБЭД().ОбменСКонтрагентами,
			Результат.ПодробноеПредставлениеОшибки, ТекстСообщения);
			
		Возврат;
	КонецЕсли;
	
	Если ПолучитьИзВременногоХранилища(Результат.АдресРезультата) = Истина Тогда
		
		Оповестить("ОбновитьСостояниеЭД");
		Оповестить("ОбновитьСостояниеПриглашений");
		Закрыть();
		
	Иначе
		
		ТекстСообщения = НСтр("ru = 'Во время удаления настройки интеркампани произошла ошибка.
                               |Подробнее см. в журнале регистрации.'");
		
		ОбщегоНазначенияКлиент.СообщитьПользователю(ТекстСообщения);
	КонецЕсли;
	
	
КонецПроцедуры

&НаКлиенте
Процедура ПередЗакрытиемОкнаФормы(Результат, ДополнительныеПараметры) Экспорт
	
	Если Результат = КодВозвратаДиалога.Отмена Тогда
		Возврат;
	ИначеЕсли Результат = КодВозвратаДиалога.Нет Тогда
		Модифицированность = Ложь;
	ИначеЕсли Результат = КодВозвратаДиалога.Да Тогда
		ЗаписатьИЗакрыть(Элементы.ФормаЗаписатьИЗакрыть);
		Возврат;
	КонецЕсли;
	
	Закрыть();
	
КонецПроцедуры

#Область КонтекстныеПодсказки

&НаСервере
Процедура СформироватьКонтекст(КатегорииПересчета = Неопределено)
	
	Если Не КонтекстныеПодсказкиБЭД.ФункционалКонтекстныхПодсказокДоступен() Тогда 
		Возврат;
	КонецЕсли;

	КонтекстныеПодсказкиБЭД.ОтобразитьАктуальныеДляКонтекстаНовости(ЭтотОбъект);
	
КонецПроцедуры

// Процедура показывает новости, требующие прочтения (важные и очень важные).
//
// Параметры:
//  Нет.
//
&НаКлиенте
Процедура Подключаемый_ПоказатьНовостиТребующиеПрочтенияПриОткрытии()

	ИдентификаторыСобытийПриОткрытии = "ПриОткрытии";	
	КонтекстныеПодсказкиБЭДКлиент.КонтекстныеНовости_ПоказатьНовостиТребующиеПрочтенияПриОткрытии(ЭтотОбъект, ИдентификаторыСобытийПриОткрытии);

КонецПроцедуры

&НаКлиенте
Процедура Подключаемый_ПанельКонтекстныхНовостей_ЭлементУправленияНажатие(Элемент)
	
	КонтекстныеПодсказкиБЭДКлиент.ЭлементУправленияНажатие(ЭтотОбъект, Элемент);
	
КонецПроцедуры

&НаКлиенте
Процедура Подключаемый_ПанельКонтекстныхНовостейОбработкаНавигационнойСсылки(Элемент, ПараметрНавигационнаяСсылка, СтандартнаяОбработка)

	СтандартнаяОбработка = Ложь;

	КонтекстныеПодсказкиБЭДКлиент.ПанельКонтекстныхНовостей_ЭлементПанелиНовостейОбработкаНавигационнойСсылки(
		ЭтотОбъект,
		Элемент,
		ПараметрНавигационнаяСсылка,
		СтандартнаяОбработка);

КонецПроцедуры

#КонецОбласти

#КонецОбласти