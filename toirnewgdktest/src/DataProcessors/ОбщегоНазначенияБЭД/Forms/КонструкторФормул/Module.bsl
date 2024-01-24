
#Область ОбработчикиСобытийФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	
	Формула         = Параметры.Формула;
	ИсходнаяФормула = Параметры.Формула;
	
	ЗагрузитьНаборОперандов(Параметры.АдресНабораОперандов);
	
	ПостроитьДеревоОператоров();
	
	УстановитьУсловноеОформление();
	
	ОбщегоНазначенияБЭД.ПроверитьФормулу(Формула, КомпоновщикНастроек.Настройки.ДоступныеПоляОтбора);
	
КонецПроцедуры

&НаКлиенте
Процедура ПередЗакрытием(Отказ, СтандартнаяОбработка)
	
	Если ПринудительноЗакрытьФорму Тогда
		Возврат;
	КонецЕсли;
	
	Если Модифицированность
		И ЗначениеЗаполнено(ИсходнаяФормула) И ИсходнаяФормула <> Формула Тогда
		
		Отказ = Истина;
		ПоказатьВопрос(
			Новый ОписаниеОповещения("ПередЗакрытиемЗавершение", ЭтотОбъект),
			НСтр("ru='Данные были изменены. Сохранить изменения?'"), РежимДиалогаВопрос.ДаНетОтмена);
		
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ПередЗакрытиемЗавершение(РезультатВопроса, ДополнительныеПараметры) Экспорт
	
	Если РезультатВопроса = КодВозвратаДиалога.Да Тогда
		
		Если ПроверитьФормулу(Формула) Тогда
			
			ПринудительноЗакрытьФорму = Истина;
			Закрыть(Формула);
			
		КонецЕсли;
		
	ИначеЕсли РезультатВопроса = КодВозвратаДиалога.Нет Тогда
		
		ПринудительноЗакрытьФорму = Истина;
		Закрыть(Неопределено);
		
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовТаблицыФормыКомпоновщикНастроек

&НаКлиенте
Процедура КомпоновщикНастроекНастройкиВыборДоступныеПоляВыбораВыбор(Элемент, ВыбраннаяСтрока, Поле, СтандартнаяОбработка)
	
	ТекстСтроки = Строка(КомпоновщикНастроек.Настройки.ДоступныеПоляПорядка.ПолучитьОбъектПоИдентификатору(ВыбраннаяСтрока).Поле);
	Операнд = ТекстОперанда(ТекстСтроки);
	Если Не ЗначениеЗаполнено(Операнд) Тогда
		ОбщегоНазначенияКлиент.СообщитьПользователю(
			НСтр("ru='Имя реквизита не должно содержать ""."". Переименуйте реквизит.'"));
		Возврат;
	КонецЕсли;
	ВставитьТекстВФормулу(Операнд);
	
КонецПроцедуры

&НаКлиенте
Процедура КомпоновщикНастроекНачалоПеретаскивания(Элемент, ПараметрыПеретаскивания, Выполнение)
	
	ТекстЭлемента = Строка(КомпоновщикНастроек.Настройки.ДоступныеПоляПорядка.ПолучитьОбъектПоИдентификатору(Элементы.КомпоновщикНастроек.ТекущаяСтрока).Поле);
	Операнд = ТекстОперанда(ТекстЭлемента);
	Если Не ЗначениеЗаполнено(Операнд) Тогда
		ОбщегоНазначенияКлиент.СообщитьПользователю(
			НСтр("ru='Имя реквизита не должно содержать ""."". Переименуйте реквизит.'"));
		Выполнение = Ложь;
		Возврат;
	КонецЕсли;
	ПараметрыПеретаскивания.Значение = Операнд;
	
КонецПроцедуры

&НаКлиенте
Процедура КомпоновщикНастроекПриАктивизацииСтроки(Элемент)
	
	ТекущиеДанные = Элементы.КомпоновщикНастроек.ТекущиеДанные;
	ПараметрыОтбора = Новый Структура("Идентификатор", ТекущиеДанные.Заголовок);
	НайденныеСтроки = ОграниченияОперандов.НайтиСтроки(ПараметрыОтбора);
	Если НайденныеСтроки.Количество() = 0 Тогда
		Элементы.ДекорацияПояснение.Заголовок = "";
		Возврат;
	КонецЕсли;
	
	МассивСтрок = Новый Массив;
	МассивСтрок.Добавить(НСтр("ru = 'Значение реквизита недоступно для источников:'"));
	Для Каждого СтрокаТаблицы Из НайденныеСтроки Цикл
		МассивСтрок.Добавить(СтрокаТаблицы.Источник);
		МассивСтрок.Добавить(";");
	КонецЦикла;
	
	МассивСтрок.Удалить(МассивСтрок.Количество() - 1);
	
	Элементы.ДекорацияПояснение.Заголовок = СтрСоединить(МассивСтрок, " ");
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовТаблицыФормыОператоры

&НаКлиенте
Процедура ОператорыВыбор(Элемент, ВыбраннаяСтрока, Поле, СтандартнаяОбработка)
	
	СтандартнаяОбработка = Ложь;
	ВставитьОператорВФормулу();
	
КонецПроцедуры

&НаКлиенте
Процедура ОператорыНачалоПеретаскивания(Элемент, ПараметрыПеретаскивания, СтандартнаяОбработка)
	
	Если ЗначениеЗаполнено(Элемент.ТекущиеДанные.Оператор) Тогда
		ПараметрыПеретаскивания.Значение = Элемент.ТекущиеДанные.Оператор;
	Иначе
		СтандартнаяОбработка = Ложь;
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ОператорыОкончаниеПеретаскивания(Элемент, ПараметрыПеретаскивания, СтандартнаяОбработка)
	
	Если Элемент.ТекущиеДанные.Оператор = "Формат(,)" Тогда
		ФорматСтроки = Новый КонструкторФорматнойСтроки;
		ФорматСтроки.Показать(Новый ОписаниеОповещения("ОператорыОкончаниеПеретаскиванияЗавершение", ЭтотОбъект, Новый Структура("ФорматСтроки", ФорматСтроки)));
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ОператорыОкончаниеПеретаскиванияЗавершение(Текст, ДополнительныеПараметры) Экспорт
    
    ФорматСтроки = ДополнительныеПараметры.ФорматСтроки;
    
    
    Если ЗначениеЗаполнено(ФорматСтроки.Текст) Тогда
        ТекстДляВставки = "Формат( , """ + ФорматСтроки.Текст + """)";
        Элементы.Формула.ВыделенныйТекст = ТекстДляВставки;
    КонецЕсли;

КонецПроцедуры

#КонецОбласти

#Область ОбработчикиКомандФормы

&НаКлиенте
Процедура ОК(Команда)
	
	ОчиститьСообщения();
	
	Если ПроверитьФормулу(Формула) Тогда
		
		ПринудительноЗакрытьФорму = Истина;
		Закрыть(Формула);
		
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура Проверить(Команда)
	
	ОчиститьСообщения();
	ПроверитьФормулуИнтерактивно(Формула);
	
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

&НаСервере
Процедура ПостроитьДеревоОператоров()
	
	Дерево = ПолучитьПустоеДеревоОператоров();
	
	ГруппаОператоров = ДобавитьГруппуОператоров(Дерево, НСтр("ru='Разделители'"));
	
	ДобавитьОператор(Дерево, "/",                 ГруппаОператоров, " + ""/"" + ");
	ДобавитьОператор(Дерево, "\",                 ГруппаОператоров, " + ""\"" + ");
	ДобавитьОператор(Дерево, "|",                 ГруппаОператоров, " + ""|"" + ");
	ДобавитьОператор(Дерево, "_",                 ГруппаОператоров, " + ""_"" + ");
	ДобавитьОператор(Дерево, ",",                 ГруппаОператоров, " + "", "" + ");
	ДобавитьОператор(Дерево, ".",                 ГруппаОператоров, " + "". "" + ");
	ДобавитьОператор(Дерево, НСтр("ru='Пробел'"), ГруппаОператоров, " + "" "" + ");
	ДобавитьОператор(Дерево, "(",                 ГруппаОператоров, " + "" ("" + ");
	ДобавитьОператор(Дерево, ")",                 ГруппаОператоров, " + "") "" + ");
	ДобавитьОператор(Дерево, """",                ГруппаОператоров, " + """""""" + ");
	
	ГруппаОператоров = ДобавитьГруппуОператоров(Дерево, НСтр("ru='Операторы'"));
	
	ДобавитьОператор(Дерево, "+", ГруппаОператоров, " + ");
	ДобавитьОператор(Дерево, "-", ГруппаОператоров, " - ");
	ДобавитьОператор(Дерево, "*", ГруппаОператоров, " * ");
	ДобавитьОператор(Дерево, "/", ГруппаОператоров, " / ");
	
	ГруппаОператоров = ДобавитьГруппуОператоров(Дерево, НСтр("ru='Логические операторы и константы'"));
	ДобавитьОператор(Дерево, "<",                 ГруппаОператоров, " < ");
	ДобавитьОператор(Дерево, ">",                 ГруппаОператоров, " > ");
	ДобавитьОператор(Дерево, "<=",                ГруппаОператоров, " <= ");
	ДобавитьОператор(Дерево, ">=",                ГруппаОператоров, " >= ");
	ДобавитьОператор(Дерево, "=",                 ГруппаОператоров, " = ");
	ДобавитьОператор(Дерево, "<>",                ГруппаОператоров, " <> ");
	ДобавитьОператор(Дерево, НСтр("ru='И'"),      ГруппаОператоров, " " + НСтр("ru='И'")      + " ");
	ДобавитьОператор(Дерево, НСтр("ru='Или'"),    ГруппаОператоров, " " + НСтр("ru='Или'")    + " ");
	ДобавитьОператор(Дерево, НСтр("ru='Не'"),     ГруппаОператоров, " " + НСтр("ru='Не'")     + " ");
	ДобавитьОператор(Дерево, НСтр("ru='ИСТИНА'"), ГруппаОператоров, " " + НСтр("ru='ИСТИНА'") + " ");
	ДобавитьОператор(Дерево, НСтр("ru='ЛОЖЬ'"),   ГруппаОператоров, " " + НСтр("ru='ЛОЖЬ'")   + " ");
	
	ГруппаОператоров = ДобавитьГруппуОператоров(Дерево, НСтр("ru='Числовые функции'"));
	
	ДобавитьОператор(Дерево, НСтр("ru='Максимум'"),    ГруппаОператоров, НСтр("ru='Макс(,)'"), 2);
	ДобавитьОператор(Дерево, НСтр("ru='Минимум'"),     ГруппаОператоров, НСтр("ru='Мин(,)'"),  2);
	ДобавитьОператор(Дерево, НСтр("ru='Округление'"),  ГруппаОператоров, НСтр("ru='Окр(,)'"),  2);
	ДобавитьОператор(Дерево, НСтр("ru='Целая часть'"), ГруппаОператоров, НСтр("ru='Цел()'"),   1);
	
	ГруппаОператоров = ДобавитьГруппуОператоров(Дерево, НСтр("ru='Строковые функции'"));
	
	ДобавитьОператор(Дерево, НСтр("ru='Строка'"),      ГруппаОператоров, НСтр("ru='Строка()'"));
	ДобавитьОператор(Дерево, НСтр("ru='ВРег'"),        ГруппаОператоров, НСтр("ru='ВРег()'"));
	ДобавитьОператор(Дерево, НСтр("ru='Лев'"),         ГруппаОператоров, НСтр("ru='Лев()'"));
	ДобавитьОператор(Дерево, НСтр("ru='НРег'"),        ГруппаОператоров, НСтр("ru='НРег()'"));
	ДобавитьОператор(Дерево, НСтр("ru='Прав'"),        ГруппаОператоров, НСтр("ru='Прав()'"));
	ДобавитьОператор(Дерево, НСтр("ru='СокрЛ'"),       ГруппаОператоров, НСтр("ru='СокрЛ()'"));
	ДобавитьОператор(Дерево, НСтр("ru='СокрЛП'"),      ГруппаОператоров, НСтр("ru='СокрЛП()'"));
	ДобавитьОператор(Дерево, НСтр("ru='СокрП'"),       ГруппаОператоров, НСтр("ru='СокрП()'"));
	ДобавитьОператор(Дерево, НСтр("ru='ТРег'"),        ГруппаОператоров, НСтр("ru='ТРег()'"));
	ДобавитьОператор(Дерево, НСтр("ru='СтрЗаменить'"), ГруппаОператоров, НСтр("ru='СтрЗаменить(,,)'"));
	ДобавитьОператор(Дерево, НСтр("ru='СтрДлина'"),    ГруппаОператоров, НСтр("ru='СтрДлина()'"));
	
	ГруппаОператоров = ДобавитьГруппуОператоров(Дерево, НСтр("ru='Прочие функции'"));
	
	ДобавитьОператор(Дерево, НСтр("ru='Условие'"),                   ГруппаОператоров, "?(,,)", 3);
	ДобавитьОператор(Дерево, НСтр("ru='Предопределенное значение'"), ГруппаОператоров, НСтр("ru='ПредопределенноеЗначение()'"));
	ДобавитьОператор(Дерево, НСтр("ru='Значение заполнено'"),        ГруппаОператоров, НСтр("ru='ЗначениеЗаполнено()'"));
	ДобавитьОператор(Дерево, НСтр("ru='Формат'"),                    ГруппаОператоров, НСтр("ru='Формат(,)'"));
	
	ЗначениеВРеквизитФормы(Дерево, "Операторы");
	
КонецПроцедуры

&НаСервере
Функция ПолучитьПустоеДеревоОператоров()
	
	Дерево = Новый ДеревоЗначений();
	Дерево.Колонки.Добавить("Наименование");
	Дерево.Колонки.Добавить("Оператор");
	Дерево.Колонки.Добавить("Сдвиг", Новый ОписаниеТипов("Число"));
	
	Возврат Дерево;
	
КонецФункции

&НаСервере
Функция ДобавитьГруппуОператоров(Дерево, Наименование)
	
	НоваяГруппа = Дерево.Строки.Добавить();
	НоваяГруппа.Наименование = Наименование;
	
	Возврат НоваяГруппа;
	
КонецФункции

&НаСервере
Функция ДобавитьОператор(Дерево, Наименование, Родитель = Неопределено, Оператор = Неопределено, Сдвиг = 0)
	
	НоваяСтрока = ?(Родитель <> Неопределено, Родитель.Строки.Добавить(), Дерево.Строки.Добавить());
	НоваяСтрока.Наименование = Наименование;
	НоваяСтрока.Оператор = ?(ЗначениеЗаполнено(Оператор), Оператор, Наименование);
	НоваяСтрока.Сдвиг = Сдвиг;
	
	Возврат НоваяСтрока;
	
КонецФункции

&НаСервере
Функция ПроверитьФормулу(Формула)
	Перем ТекстОшибки;
	
	Результат = ОбщегоНазначенияБЭД.ПроверитьФормулу(Формула, КомпоновщикНастроек.Настройки.ДоступныеПоляОтбора, ТекстОшибки);
	
	Если ЗначениеЗаполнено(ТекстОшибки) Тогда
		ОбщегоНазначения.СообщитьПользователю(ТекстОшибки,, "Формула");
	КонецЕсли;
	
	Возврат Результат;
	
КонецФункции 

&НаКлиенте
Процедура ПроверитьФормулуИнтерактивно(Формула)
	
	Если ЗначениеЗаполнено(Формула) Тогда
		Если ПроверитьФормулу(Формула) Тогда
			ПоказатьОповещениеПользователя(НСтр("ru = 'В формуле ошибок не обнаружено'"),,,БиблиотекаКартинок.Информация32);
		КонецЕсли;
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ВставитьТекстВФормулу(ТекстДляВставки, Сдвиг = 0)
	
	СтрокаНач = 0;
	СтрокаКон = 0;
	КолонкаНач = 0;
	КолонкаКон = 0;
	
	Элементы.Формула.ПолучитьГраницыВыделения(СтрокаНач, КолонкаНач, СтрокаКон, КолонкаКон);
	
	Если (КолонкаКон = КолонкаНач) И (КолонкаКон + СтрДлина(ТекстДляВставки)) > Элементы.Формула.Ширина / 8 Тогда
		Элементы.Формула.ВыделенныйТекст = "";
	КонецЕсли;
		
	Элементы.Формула.ВыделенныйТекст = ТекстДляВставки;
	
	Если Не Сдвиг = 0 Тогда
		Элементы.Формула.ПолучитьГраницыВыделения(СтрокаНач, КолонкаНач, СтрокаКон, КолонкаКон);
		Элементы.Формула.УстановитьГраницыВыделения(СтрокаНач, КолонкаНач - Сдвиг, СтрокаКон, КолонкаКон - Сдвиг);
	КонецЕсли;
		
	ТекущийЭлемент = Элементы.Формула;
	
КонецПроцедуры

&НаСервере
Процедура ЗагрузитьНаборОперандов(АдресНабораОперандов)
	
	Отказ = Ложь;
	
	Если НЕ ЗначениеЗаполнено(АдресНабораОперандов) Тогда
		Отказ = Истина;
	Иначе
		ТекстЗапроса = ПолучитьИзВременногоХранилища(АдресНабораОперандов);
		Если НЕ ЗначениеЗаполнено(ТекстЗапроса) Тогда
			Отказ = Истина;
		КонецЕсли;
	КонецЕсли;
	
	Если Отказ Тогда
		Элементы.ГруппаОперанды.Видимость = Ложь;
		Возврат;
	КонецЕсли;
	
	ЗаполнитьТаблицуОперандов(ТекстЗапроса);
	
	ЗаполнитьОграниченияОперандов(ТекстЗапроса);
	
КонецПроцедуры

&НаСервере
Процедура ЗаполнитьТаблицуОперандов(ТекстЗапроса)
	
	СхемаКомпоновки = Новый СхемаКомпоновкиДанных;
	
	Источник = СхемаКомпоновки.ИсточникиДанных.Добавить();
	Источник.Имя = "Источник";
	Источник.ТипИсточникаДанных = "Local";
	
	Набор = СхемаКомпоновки.НаборыДанных.Добавить(Тип("НаборДанныхЗапросСхемыКомпоновкиДанных"));
	Набор.Имя = Источник.Имя;
	Набор.ИсточникДанных = Источник.Имя;
	Набор.АвтоЗаполнениеДоступныхПолей = Истина;
	Набор.Запрос = ТекстЗапроса;
	
	АдресСхемыКомпоновки = ПоместитьВоВременноеХранилище(СхемаКомпоновки, УникальныйИдентификатор);
	КомпоновщикНастроек.Инициализировать(Новый ИсточникДоступныхНастроекКомпоновкиДанных(АдресСхемыКомпоновки));
	
КонецПроцедуры

&НаСервере
Процедура ЗаполнитьОграниченияОперандов(ТекстЗапроса)
	
	СхемаЗапроса = Новый СхемаЗапроса;
	СхемаЗапроса.УстановитьТекстЗапроса(ТекстЗапроса);
	
	Если СхемаЗапроса.ПакетЗапросов[0].Операторы.Количество() < 2 Тогда
		Возврат;
	КонецЕсли;
	
	Источники = Новый Массив;
	Для Каждого ОператорСхемы Из СхемаЗапроса.ПакетЗапросов[0].Операторы Цикл
		ИмяТаблицы = ОператорСхемы.Источники[0].Источник.ИмяТаблицы;
		МетаданныеОбъекта = Метаданные.НайтиПоПолномуИмени(ИмяТаблицы);
		Источники.Добавить(МетаданныеОбъекта.Представление());
	КонецЦикла;
	
	Для Каждого КолонкаСхемы Из СхемаЗапроса.ПакетЗапросов[0].Колонки Цикл
		
		Счетчик = 0;
		Для Каждого Поле Из КолонкаСхемы.Поля Цикл
			
			Если Поле = Неопределено Тогда
				НоваяСтрока = ОграниченияОперандов.Добавить();
				НоваяСтрока.Идентификатор = КолонкаСхемы.Псевдоним;
				НоваяСтрока.Источник = Источники[Счетчик];
			КонецЕсли;
			
			Счетчик = Счетчик + 1;
		КонецЦикла;
		
	КонецЦикла;
	
КонецПроцедуры

&НаКлиенте
Процедура ВставитьОператорВФормулу()
	
	Если Элементы.Операторы.ТекущиеДанные.Наименование = "Формат" Тогда
		ФорматСтроки = Новый КонструкторФорматнойСтроки;
		ФорматСтроки.Показать(Новый ОписаниеОповещения("ВставитьОператорВФормулуЗавершение", ЭтотОбъект, Новый Структура("ФорматСтроки", ФорматСтроки)));
		Возврат;
	Иначе
		ВставитьТекстВФормулу(Элементы.Операторы.ТекущиеДанные.Оператор, Элементы.Операторы.ТекущиеДанные.Сдвиг);
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ВставитьОператорВФормулуЗавершение(Текст, ДополнительныеПараметры) Экспорт

	ФорматСтроки = ДополнительныеПараметры.ФорматСтроки;

	Если ЗначениеЗаполнено(ФорматСтроки.Текст) Тогда
		ТекстДляВставки = "Формат( , """ + ФорматСтроки.Текст + """)";
		ВставитьТекстВФормулу(ТекстДляВставки, Элементы.Операторы.ТекущиеДанные.Сдвиг);
	Иначе
		ВставитьТекстВФормулу(Элементы.Операторы.ТекущиеДанные.Оператор, Элементы.Операторы.ТекущиеДанные.Сдвиг);
	КонецЕсли;

КонецПроцедуры

&НаКлиентеНаСервереБезКонтекста
Функция ТекстОперанда(ТекстОперанда)
	
	Возврат СтрШаблон("[%1]", ТекстОперанда);
	
КонецФункции

&НаСервере
Процедура УстановитьУсловноеОформление()
	
	УсловноеОформление.Элементы.Очистить();
	
	ЭлементОформления = УсловноеОформление.Элементы.Добавить();
	
	ПолеЭлемента = ЭлементОформления.Поля.Элементы.Добавить();
	ПолеЭлемента.Поле = Новый ПолеКомпоновкиДанных("КомпоновщикНастроекЗаголовок");
	
	Список = Новый СписокЗначений;
	Список.ЗагрузитьЗначения(ОбщегоНазначения.ВыгрузитьКолонку(
		ОграниченияОперандов, "Идентификатор", Истина));
	
	ОтборЭлемента = ЭлементОформления.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
	ОтборЭлемента.ЛевоеЗначение = Новый ПолеКомпоновкиДанных("КомпоновщикНастроек.Настройки.Отбор.ДоступныеПоляОтбора.Заголовок");
	ОтборЭлемента.ВидСравнения = ВидСравненияКомпоновкиДанных.ВСписке;
	ОтборЭлемента.ПравоеЗначение = Список;
	
	ЭлементОформления.Оформление.УстановитьЗначениеПараметра("ЦветТекста", ЦветаСтиля.ЦветГиперссылкиБЭД);
	
КонецПроцедуры

#КонецОбласти
