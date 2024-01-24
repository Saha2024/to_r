
#Область ПрограммныйИнтерфейс

#Область МетодыРаботыАвтозаполненияРеквизитовДокумента

// Служебная функция, предназначенная для получения описания типов числа, заданной разрядности.
// 
// Параметры:
//  Разрядность 			- Число - разряд числа.
//  РазрядностьДробнойЧасти - Число - разряд дробной части.
//
// Возвращаемое значение:
//  ОписаниеТипов - Объект "ОписаниеТипов" для числа указанной разрядности.
//
Функция ПолучитьОписаниеТиповЧисла(Разрядность, РазрядностьДробнойЧасти) Экспорт

	Возврат Новый ОписаниеТипов("Число", Новый КвалификаторыЧисла(Разрядность, РазрядностьДробнойЧасти));

КонецФункции // ПолучитьОписаниеТиповЧисла()

// Служебная функция, предназначенная для получения описания типов даты
// 
// Параметры:
//  ЧастиДаты - ЧастиДаты - системное перечисление ЧастиДаты.
//
// Возвращаемое значение:
//		ОписаниеТипов - описание типов даты.
Функция ПолучитьОписаниеТиповДаты(ЧастиДаты) Экспорт

	Возврат Новый ОписаниеТипов("Дата", , , Новый КвалификаторыДаты(ЧастиДаты));

КонецФункции // ПолучитьОписаниеТиповДаты()

// Заполняет колонку таблицы значений последовательными номерами
// Параметры:
//		Таблица - ТаблицаЗначений - строки которой нужно пронумеровать
//		ИмяКолонкиНомераСтроки - строка - колонка таблицы значений, в которой будут указаны номера строк.
Процедура ПронумероватьТаблицуЗначений(Таблица, ИмяКолонкиНомераСтроки) Экспорт

	Таблица.Колонки.Добавить(ИмяКолонкиНомераСтроки, ПолучитьОписаниеТиповЧисла(15, 0));

	КоличествоСтрок = Таблица.Количество() - 1;
	Для НомерСтроки = 0 По КоличествоСтрок Цикл
		Таблица[НомерСтроки][ИмяКолонкиНомераСтроки] = НомерСтроки;
	КонецЦикла;

КонецПроцедуры

// Возвращает ключ данных для подстановки в сообщение пользователю
// Ключ данных нужен при групповой обработке объектов - если он установлен, то при нажатии пользователем на сообщение
// будет открываться форма объекта.
//
//	Параметры:
//		Объект - Произвольный - объект, для которого нужно получить ключ данных.
//	
//	Возвращаемое значение:
//		Ссылка - ссылка на объект информационной базы.
//
Функция КлючДанныхДляСообщенияПользователю(Объект) Экспорт
	
	КлючДанных = Неопределено;
	XMLТипЗнч = XMLТипЗнч(Объект); 
	
	Если XMLТипЗнч <> Неопределено Тогда
		ТипЗначенияСтрокой = XMLТипЗнч.ИмяТипа;
		Если Найти(ТипЗначенияСтрокой, "Object.") > 0 Тогда
			КлючДанных = Объект.Ссылка;
		КонецЕсли;
	КонецЕсли;
	
	Возврат КлючДанных;
	
КонецФункции

#КонецОбласти

#Область ПрочиеПроцедурыИФункции

// Осуществляет проверку заполненности проверяемых реквизитов.
//
// Параметры:
// 		Объект - ДокументОбъект, СправочникОбъект - Проверяемый объект.
// 		МассивПроверяемыхРеквизитов - Массив из Строка - массив проверяемых реквизитов.
//
// Возвращаемое значение:
// 		Булево - Истина, если значение хотя бы одного реквизита не заполнено, иначе Ложь.
//
Функция ПроверитьЗаполнениеРеквизитовОбъекта(Объект, МассивПроверяемыхРеквизитов) Экспорт
	
	Перем ПроверяемыеРеквизитыТЧ;
	Отказ = Ложь;
	
	// Получение метаданных объекта
	МетаданныеОбъекта = Объект.Ссылка.Метаданные();
	
	// Создание структуры стандартных реквизитов
	СтандартныеРеквизиты = Новый Структура;
	Для Каждого Реквизит Из МетаданныеОбъекта.СтандартныеРеквизиты Цикл
		СтандартныеРеквизиты.Вставить(Реквизит.Имя, ?(ЗначениеЗаполнено(Реквизит.Синоним), Реквизит.Синоним, Реквизит.Имя));
	КонецЦикла;
	
	// Создание структуры для хранения имен табличных частей и проверяемых реквизитов в них.
	// 		Ключ -  Имя табличной части
	// 		Значение - Массив - Массив строк, реквизитов этой табличной части для проверки.
	ТабличныеЧасти = Новый Структура;
	
	// Создание шаблонов сообщений об ошибках не заполненных реквизитов и реквизитов табличных частей.
	ШаблонОшибкиРеквизита = НСтр("ru='Поле ""%ИмяРеквизита%"" не заполнено'");
	ШаблонОшибкиТЧ = НСтр("ru='Не введено ни одной строки в список ""%ИмяРеквизита%""'");
	ШаблонОшибкиРеквизитаТЧ = НСтр("ru='Не заполнена колонка ""%ИмяРеквизита%"" в строке %НомерСтроки% списка ""%ИмяТабличнойЧасти%""'");
	
	// Проверка реквизитов объекта и заполнение структуры по реквизитам табличных частей.
	Для Каждого Реквизит Из МассивПроверяемыхРеквизитов Цикл
		
		ПозицияТочки = Найти(Реквизит,".");
		
		Если ПозицияТочки > 0 Тогда // В случае если указан реквизит табличной части
			
			ДлинаСтроки       = СтрДлина(Реквизит);
			ИмяТабличнойЧасти = Лев(Реквизит, ПозицияТочки-1);
			ИмяРеквизита      = Прав(Реквизит, ДлинаСтроки - ПозицияТочки);
			
			// Сохранение проверяемого реквизита табличной части в структуру
			Если НЕ ТабличныеЧасти.Свойство(ИмяТабличнойЧасти, ПроверяемыеРеквизитыТЧ) Тогда
				ПроверяемыеРеквизитыТЧ = Новый Массив;
				ТабличныеЧасти.Вставить(ИмяТабличнойЧасти, ПроверяемыеРеквизитыТЧ);
			КонецЕсли;
			ПроверяемыеРеквизитыТЧ.Добавить(ИмяРеквизита);
			
		Иначе // В случае если указан реквизит объекта
			
			Если Не ЗначениеЗаполнено(Объект[Реквизит]) Тогда
				
				Если МетаданныеОбъекта.Реквизиты.Найти(Реквизит) <> Неопределено Тогда // Если указано имя реквизита
					ТекстОшибки = СтрЗаменить(ШаблонОшибкиРеквизита, "%ИмяРеквизита%",
						МетаданныеОбъекта.Реквизиты[Реквизит].Синоним);
				ИначеЕсли СтандартныеРеквизиты.Свойство(Реквизит) Тогда // Если указано имя стандартного реквизита
					ТекстОшибки = СтрЗаменить(ШаблонОшибкиРеквизита, "%ИмяРеквизита%",
						СтандартныеРеквизиты[Реквизит]);
				Иначе // Если указано имя табличной части
					ТекстОшибки = СтрЗаменить(ШаблонОшибкиТЧ, "%ИмяРеквизита%",
						МетаданныеОбъекта.ТабличныеЧасти[Реквизит].Синоним);
				КонецЕсли;
				
				ОбщегоНазначения.СообщитьПользователю(ТекстОшибки, Объект, Реквизит,, Отказ);
				
			КонецЕсли;
			
		КонецЕсли;
		
	КонецЦикла;
	
	// Проверка реквизитов в табличных частях
	Для Каждого ТабличнаяЧасть Из ТабличныеЧасти Цикл
		
		ИмяТабличнойЧасти = ТабличнаяЧасть.Ключ;
		ТабличнаяЧастьОбъекта = Объект[ТабличнаяЧасть.Ключ];
		МассивРеквизитов = ТабличнаяЧасть.Значение;
		
		// Цикл по всем строкам табличной части.
		Для НомерСтроки=0 По ТабличнаяЧастьОбъекта.Количество()-1 Цикл
			
			// Цикл по всем проверяемым реквизитам для текущей табличной части.
			Для НомерРеквизита=0 По МассивРеквизитов.Количество()-1 Цикл
				
				ИмяРеквизита = МассивРеквизитов[НомерРеквизита];
				
				Если Не ЗначениеЗаполнено(ТабличнаяЧастьОбъекта[НомерСтроки][ИмяРеквизита]) Тогда
					
					ТекстОшибки = СтрЗаменить(ШаблонОшибкиРеквизитаТЧ, "%ИмяРеквизита%", МетаданныеОбъекта.ТабличныеЧасти[ИмяТабличнойЧасти].Реквизиты[ИмяРеквизита].Синоним);
					ТекстОшибки = СтрЗаменить(ТекстОшибки, "%НомерСтроки%", Формат(НомерСтроки+1, "ЧГ=0"));
					ТекстОшибки = СтрЗаменить(ТекстОшибки, "%ИмяТабличнойЧасти%", МетаданныеОбъекта.ТабличныеЧасти[ИмяТабличнойЧасти].Синоним);
					
					ОбщегоНазначения.СообщитьПользователю(ТекстОшибки, Объект, ОбщегоНазначенияКлиентСервер.ПутьКТабличнойЧасти(ИмяТабличнойЧасти, НомерСтроки+1, ИмяРеквизита),, Отказ);
					
				КонецЕсли;
				
			КонецЦикла;
			
		КонецЦикла;
		
	КонецЦикла;
	
	МассивПроверяемыхРеквизитов.Очистить();
	
	Возврат Отказ;
	
КонецФункции // ПроверитьЗаполнениеРеквизитовОбъекта()

// Функция возвращает список разрешенных для чтения пользователей, принадлежащих ролям указанным в качестве параметра.
//
// Параметры:
//			МассивРолей - Массив из Строка - Массив со строковыми именами ролей.
//			ПредставлениеТекущегоПользователя - Строка - представление пользователя.
//
// Возвращаемое значение:
//		СписокЗначений - список пользователей.
Функция ПолучитьСписокПользователейПоМассивуРолей(МассивРолей, ПредставлениеТекущегоПользователя = Неопределено) Экспорт
	
	СписокПользователей = Новый СписокЗначений;
	Если ПредставлениеТекущегоПользователя = Неопределено Тогда
		ПредставлениеТекущегоПользователя = НСтр("ru='<Мои документы>'");
	КонецЕсли;
	
	// Запрос получения списка всех пользователей. Выполняется в привилегированном режиме.
	УстановитьПривилегированныйРежим(Истина);
	
	Запрос = Новый Запрос;
	Запрос.Текст =
	"ВЫБРАТЬ
	|	ПрофилиГруппДоступаРоли.Ссылка КАК Ссылка
	|ПОМЕСТИТЬ ПрофилиГруппДоступа
	|ИЗ
	|	Справочник.ПрофилиГруппДоступа.Роли КАК ПрофилиГруппДоступаРоли
	|ГДЕ
	|	ПрофилиГруппДоступаРоли.Роль.Имя В(&МассивРолей)
	|
	|СГРУППИРОВАТЬ ПО
	|	ПрофилиГруппДоступаРоли.Ссылка
	|
	|ИНДЕКСИРОВАТЬ ПО
	|	Ссылка
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	ЕСТЬNULL(ГруппыПользователейСостав.Пользователь, ГруппыДоступаПользователи.Пользователь) КАК Пользователь
	|ИЗ
	|	Справочник.ГруппыДоступа.Пользователи КАК ГруппыДоступаПользователи
	|		ЛЕВОЕ СОЕДИНЕНИЕ Справочник.ГруппыПользователей.Состав КАК ГруппыПользователейСостав
	|		ПО ГруппыДоступаПользователи.Пользователь = ГруппыПользователейСостав.Ссылка
	|ГДЕ
	|	ГруппыДоступаПользователи.Ссылка.Профиль В
	|			(ВЫБРАТЬ
	|				ПрофилиГруппДоступа.Ссылка
	|			ИЗ
	|				ПрофилиГруппДоступа КАК ПрофилиГруппДоступа)
	|
	|СГРУППИРОВАТЬ ПО
	|	ЕСТЬNULL(ГруппыПользователейСостав.Пользователь, ГруппыДоступаПользователи.Пользователь)";
	
	Запрос.УстановитьПараметр("МассивРолей", МассивРолей);
	
	МассивВсехПользователей = Запрос.Выполнить().Выгрузить().ВыгрузитьКолонку("Пользователь");
	
	УстановитьПривилегированныйРежим(Ложь);
	
	// Запрос получения списка "разрешенных" пользователей
	Запрос.Текст =
	"ВЫБРАТЬ РАЗРЕШЕННЫЕ
	|	Пользователи.Ссылка КАК Пользователь
	|ИЗ
	|	Справочник.Пользователи КАК Пользователи
	|ГДЕ
	|	Пользователи.Ссылка В(&МассивПользователей)
	|	И Пользователи.ПометкаУдаления = ЛОЖЬ
	|	И Пользователи.Недействителен = ЛОЖЬ
	|
	|УПОРЯДОЧИТЬ ПО
	|	Пользователи.Наименование";
	Запрос.УстановитьПараметр("МассивПользователей", МассивВсехПользователей);
	
	ТекущийПользователь = Пользователи.ТекущийПользователь();
	
	Выборка = Запрос.Выполнить().Выбрать();
	Пока Выборка.Следующий() Цикл
		Если ТекущийПользователь = Выборка.Пользователь Тогда
			СписокПользователей.Вставить(0, Выборка.Пользователь, ПредставлениеТекущегоПользователя);
		Иначе
			СписокПользователей.Добавить(Выборка.Пользователь);
		КонецЕсли;
	КонецЦикла;
	
	Возврат СписокПользователей;
	
КонецФункции // ПолучитьСписокПользователейПоМассивуРолей()

// Возвращает разницу между двумя датами (в днях)
//
// Параметры:
//  ДатаНачала 	- Дата - начальная дата периода
//  ДатаОкончания	- Дата - конечная дата периода.
//  Периодичность	- ПеречислениеСсылка.Периодичность - периодичность.
//
// Возвращаемое значение:
//   Число	 - количество дней между двумя датами.
//
Функция РазностьДат(ДатаНачала, ДатаОкончания, Периодичность) Экспорт
	
	СекундВМинуте = 60;
	МинутВЧасе = 60;
	ЧасовВДне = 24;
	
	Если Периодичность = Перечисления.Периодичность.Год Тогда
		Возврат Год(ДатаОкончания) - Год(ДатаНачала);
		
	ИначеЕсли Периодичность = Перечисления.Периодичность.Полугодие Тогда
		Возврат ?(Месяц(ДатаОкончания)>6, 2, 1) - ?(Месяц(ДатаНачала)>6, 2, 1) + 2*(Год(ДатаОкончания) - Год(ДатаНачала));
		
	ИначеЕсли Периодичность = Перечисления.Периодичность.Квартал Тогда
		Возврат Цел(Месяц(НачалоКвартала(ДатаОкончания))/3) - Цел(Месяц(НачалоКвартала(ДатаНачала))/3) + 4*(Год(ДатаОкончания) - Год(ДатаНачала));
		
	ИначеЕсли Периодичность = Перечисления.Периодичность.Месяц Тогда
		Возврат Месяц(ДатаОкончания) - Месяц(ДатаНачала) + 12*(Год(ДатаОкончания) - Год(ДатаНачала));
		
	ИначеЕсли Периодичность = Перечисления.Периодичность.Декада Тогда
		Возврат Цел((ДатаОкончания - ДатаНачала)/(10 * СекундВМинуте*МинутВЧасе*ЧасовВДне));
		
	ИначеЕсли Периодичность = Перечисления.Периодичность.Неделя Тогда
		Возврат Цел((НачалоНедели(ДатаОкончания) - НачалоНедели(ДатаНачала))/(7 * СекундВМинуте*МинутВЧасе*ЧасовВДне));
		
	ИначеЕсли Периодичность = Перечисления.Периодичность.День Тогда
		Возврат (ДатаОкончания - ДатаНачала)/(СекундВМинуте*МинутВЧасе*ЧасовВДне);
		
	КонецЕсли;
	
КонецФункции

#КонецОбласти

#Область НастройкиФорм

// Функция возвращает представление клавиши
// Параметры:
//		ЗначениеКлавиша - Клавиша - клавиша.
//
// Возвращаемое значение:
//		Строка - Представление клавиши.
//
Функция ПредставлениеКлавиши(ЗначениеКлавиша) Экспорт
	
	Если Строка(Клавиша._1) = Строка(ЗначениеКлавиша) Тогда
		Возврат "1";
	ИначеЕсли Строка(Клавиша._2) = Строка(ЗначениеКлавиша) Тогда
		Возврат "2";
	ИначеЕсли Строка(Клавиша._3) = Строка(ЗначениеКлавиша) Тогда
		Возврат "3";
	ИначеЕсли Строка(Клавиша._4) = Строка(ЗначениеКлавиша) Тогда
		Возврат "4";
	ИначеЕсли Строка(Клавиша._5) = Строка(ЗначениеКлавиша) Тогда
		Возврат "5";
	ИначеЕсли Строка(Клавиша._6) = Строка(ЗначениеКлавиша) Тогда
		Возврат "6";
	ИначеЕсли Строка(Клавиша._7) = Строка(ЗначениеКлавиша) Тогда
		Возврат "7";
	ИначеЕсли Строка(Клавиша._8) = Строка(ЗначениеКлавиша) Тогда
		Возврат "8";
	ИначеЕсли Строка(Клавиша._9) = Строка(ЗначениеКлавиша) Тогда
		Возврат "9";
	ИначеЕсли Строка(Клавиша.Num0) = Строка(ЗначениеКлавиша) Тогда
		Возврат "Num 0";
	ИначеЕсли Строка(Клавиша.Num1) = Строка(ЗначениеКлавиша) Тогда
		Возврат "Num 1";
	ИначеЕсли Строка(Клавиша.Num2) = Строка(ЗначениеКлавиша) Тогда
		Возврат "Num 2";
	ИначеЕсли Строка(Клавиша.Num3) = Строка(ЗначениеКлавиша) Тогда
		Возврат "Num 3";
	ИначеЕсли Строка(Клавиша.Num4) = Строка(ЗначениеКлавиша) Тогда
		Возврат "Num 4";
	ИначеЕсли Строка(Клавиша.Num5) = Строка(ЗначениеКлавиша) Тогда
		Возврат "Num 5";
	ИначеЕсли Строка(Клавиша.Num6) = Строка(ЗначениеКлавиша) Тогда
		Возврат "Num 6";
	ИначеЕсли Строка(Клавиша.Num7) = Строка(ЗначениеКлавиша) Тогда
		Возврат "Num 7";
	ИначеЕсли Строка(Клавиша.Num8) = Строка(ЗначениеКлавиша) Тогда
		Возврат "Num 8";
	ИначеЕсли Строка(Клавиша.Num9) = Строка(ЗначениеКлавиша) Тогда
		Возврат "Num 9";
	ИначеЕсли Строка(Клавиша.NumAdd) = Строка(ЗначениеКлавиша) Тогда
		Возврат "Num +";
	ИначеЕсли Строка(Клавиша.NumDecimal) = Строка(ЗначениеКлавиша) Тогда
		Возврат "Num .";
	ИначеЕсли Строка(Клавиша.NumDivide) = Строка(ЗначениеКлавиша) Тогда
		Возврат "Num /";
	ИначеЕсли Строка(Клавиша.NumMultiply) = Строка(ЗначениеКлавиша) Тогда
		Возврат "Num *";
	ИначеЕсли Строка(Клавиша.NumSubtract) = Строка(ЗначениеКлавиша) Тогда
		Возврат "Num -";
	Иначе
		Возврат Строка(ЗначениеКлавиша);
	КонецЕсли;
	
КонецФункции

// Функция возвращает представление клавиши
// Параметры:
//		СочетаниеКлавиш - СочетаниеКлавиш - Сочетание клавиш для которого нужно сформировать представление
//		БезСкобок - Булево - Флаг, указывающий, что представление должно быть сформировано без скобок.
//
// Возвращаемое значение:
//		Строка - Представление сочетания клавиш.
//
Функция ПредставлениеСочетанияКлавиш(СочетаниеКлавиш, БезСкобок = Ложь) Экспорт
	
	Если СочетаниеКлавиш.Клавиша = Клавиша.Нет Тогда
		Возврат "";
	КонецЕсли;
	
	Наименование = ?(БезСкобок, "", "(");
	Если СочетаниеКлавиш.Ctrl Тогда
		Наименование = Наименование + "Ctrl+"
	КонецЕсли;
	Если СочетаниеКлавиш.Alt Тогда
		Наименование = Наименование + "Alt+"
	КонецЕсли;
	Если СочетаниеКлавиш.Shift Тогда
		Наименование = Наименование + "Shift+"
	КонецЕсли;
	Наименование = Наименование + ПредставлениеКлавиши(СочетаниеКлавиш.Клавиша) + ?(БезСкобок, "", ")");
	
	Возврат Наименование;
	
КонецФункции

// Инициализирует параметры формф для полнотекствого поиска.
// Параметры:
//		Форма - ФормаКлиентскогоПриложения - форма.
//		ИмяФОИспользованияППД - Строка - имя функциональной опции.
Процедура ИнициализироватьРеквизитыФормыДляПолнотекстовогоПоиска(Форма, ИмяФОИспользованияППД) Экспорт
	
	Форма.ИнформационнаяБазаФайловая      = ОбщегоНазначения.ИнформационнаяБазаФайловая();
	Форма.ИспользоватьПолнотекстовыйПоиск = ОбщегоНазначенияУТВызовСервера.ИспользуетсяПолнотекстовыйПоиск(ИмяФОИспользованияППД);
	
	Если Форма.ИспользоватьПолнотекстовыйПоиск Тогда
			
		Форма.ИндексПолнотекстовогоПоискаАктуален = ПолнотекстовыйПоискСервер.ИндексПоискаАктуален();
		
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти

#Область ЗаполненияТабличныхЧастейДокумента

// Функция возвращает параметры проверки заполнения количества
//
//	Возвращаемое значение:
//		Структура - с полями:
//			*ИмяТЧ - Строка - значение по умолчанию "Товары"
//			*СуффиксДопРеквизита - Строка - значение по умолчанию "" - если в ТЧ два реквизита
//											"Количество", то второй назван с суффиком. 
//											если суффикс передан, то проверяются оба реквизита.
//
Функция ПараметрыПроверкиЗаполненияКоличества() Экспорт
	
	ПараметрыПроверки = Новый Структура;
	ПараметрыПроверки.Вставить("ИмяТЧ",                 "Товары");	
	ПараметрыПроверки.Вставить("СуффиксДопРеквизита",   "");
	
	Возврат ПараметрыПроверки;
	
КонецФункции

// Процедуры проверки заполнения реквизита Количество и КоличествоУпаковок в документах.
// Параметры:
//		Объект - ДокументОбъект - проверяемый ДокументОбъект
//		ПроверяемыеРеквизиты - Массив из Строка - массив проверяемых реквизитов
//		Отказ - Булево - отказ продолжения операции.
//		ПараметрыПроверки - Структура - структура параметров.
Процедура ПроверитьЗаполнениеКоличества(Объект, ПроверяемыеРеквизиты, Отказ,ПараметрыПроверки = Неопределено) Экспорт
	
	Перем ИмяТЧ;
	Перем ЗаполнятьРеквизитОбязательно;
	
	Если ПараметрыПроверки <> Неопределено Тогда
		ПараметрыПроверки.Свойство("ИмяТЧ",ИмяТЧ);
	КонецЕсли;
		
	Если Не ЗначениеЗаполнено(ИмяТЧ) Тогда
		ИмяТЧ = "Товары";	
	КонецЕсли;
	
	КлючДанных = КлючДанныхДляСообщенияПользователю(Объект);
	
	МетаданныеОбъекта = Объект.Метаданные();
	
	РеквизитПроверки =  ПроверяемыеРеквизиты.Найти(ИмяТЧ + ".КоличествоУпаковок");
	ЗаполнятьРеквизитОбязательно = РеквизитПроверки <> Неопределено;
	Если РеквизитПроверки <> Неопределено Тогда
		ПроверяемыеРеквизиты.Удалить(РеквизитПроверки);
		ПроверяемыеРеквизиты.Удалить(ПроверяемыеРеквизиты.Найти(ИмяТЧ + ".Количество"));
	КонецЕсли;
	
	ПредставлениеТЧ                          = МетаданныеОбъекта.ТабличныеЧасти[ИмяТЧ].Синоним;
	ПредставлениеРеквизитаКоличествоУпаковок = МетаданныеОбъекта.ТабличныеЧасти[ИмяТЧ].Реквизиты.КоличествоУпаковок.Синоним;
	
	ШаблонОшибкаКоличества = НСтр("ru = 'Не заполнена колонка ""%ПредставлениеКолонки%"" в строке %НомерСтроки% списка ""%ПредставлениеТЧ%""'");
	ШаблонОшибкаПересчета = НСтр("ru = 'Обнаружено нулевое количество при пересчете в единицу хранения в строке %НомерСтроки% списка ""%ПредставлениеТЧ%""'");
	
	Для Каждого СтрокаТаб Из Объект[ИмяТЧ] Цикл
		
		Если СтрокаТаб.Количество = 0 
			И СтрокаТаб.КоличествоУпаковок <> 0 Тогда
			
			ТекстСообщения = СтрЗаменить(ШаблонОшибкаПересчета, "%НомерСтроки%", Строка(СтрокаТаб.НомерСтроки));
			ТекстСообщения = СтрЗаменить(ТекстСообщения, "%ПредставлениеТЧ%", ПредставлениеТЧ);
			
			Поле = ОбщегоНазначенияКлиентСервер.ПутьКТабличнойЧасти(ИмяТЧ, СтрокаТаб.НомерСтроки, "КоличествоУпаковок");
			
			ОбщегоНазначения.СообщитьПользователю(ТекстСообщения,КлючДанных,Поле,"Объект",Отказ);
			
		ИначеЕсли ЗаполнятьРеквизитОбязательно
			И (СтрокаТаб.Количество = 0
			Или СтрокаТаб.КоличествоУпаковок = 0) Тогда
			
			ТекстСообщения = СтрЗаменить(ШаблонОшибкаКоличества, "%НомерСтроки%", Строка(СтрокаТаб.НомерСтроки));
			ТекстСообщения = СтрЗаменить(ТекстСообщения, "%ПредставлениеТЧ%", ПредставлениеТЧ);
			ТекстСообщения = СтрЗаменить(ТекстСообщения, "%ПредставлениеКолонки%", ПредставлениеРеквизитаКоличествоУпаковок);
			
			Поле = ОбщегоНазначенияКлиентСервер.ПутьКТабличнойЧасти(ИмяТЧ, СтрокаТаб.НомерСтроки, "КоличествоУпаковок");
			
			ОбщегоНазначения.СообщитьПользователю(ТекстСообщения,КлючДанных,Поле,"Объект",Отказ);
					
		КонецЕсли;
				
	КонецЦикла;
	
КонецПроцедуры

// Процедура проверяет отмену строк в табличной части документа
// Параметры:
//		Объект - ДокументОбъект - проверяемый ДокументОбъект
//		ИмяТЧ - Строка - имя проверяемой табличной части
//		ИмяРеквизитаОтменыСтрок - Строка - имя проверяемого булевого реквизита ТЧ.
//
// Возвращаемое значение:
//		Булево - все строки отменены.
Функция ВсеСтрокиОтменены(Объект, ИмяТЧ, ИмяРеквизитаОтменыСтрок) Экспорт
	
	СтруктураОтбора = Новый Структура(ИмяРеквизитаОтменыСтрок, Ложь);
	НеОтмененныеСтроки = Объект[ИмяТЧ].НайтиСтроки(СтруктураОтбора);
	
	Возврат (НеОтмененныеСтроки.Количество() = 0);
	
КонецФункции

#КонецОбласти

#КонецОбласти
