# ScriptsLibrary Отчет сравнения составов
Отчет представляет собой скрипт работающий на базе PDM-PLM системы IPS. Для скриптов используется стандарт c# 5.
Отчет показывает разницу в составах сборочных единиц по извещению и базовой версии. Извещение - документ, к которому прикрепляются новые объекты.

Сборочная единица: 
![image](https://user-images.githubusercontent.com/44323790/142209412-05cc40b8-760c-4ce1-ba1f-81160db8585a.png)

Извещение:
![image](https://user-images.githubusercontent.com/44323790/142209601-dc9da7a1-1ea1-4b28-a29f-60235c232232.png)

После того как извещение подтверждают, объекты добавляются в состав базовой версии сборки.
Отчеты запускаются на объектах типа «Сборочная единица» с включенным контекстом редактирования, т.е. в поле «Контекст редактирования» на панели инструментов должно быть выбрано технологическое (тип объекта «Извещение об изменении технологическое») или конструкторское (тип объекта «Извещение об изменении конструкторское») извещение об изменении.
В отчет выводятся объекты, в которых произошло изменение после извещения.

Пример фрагмента отчета:
![image](https://user-images.githubusercontent.com/44323790/142163876-03e2bbc3-aed6-4f04-84ee-685a56f13bc9.png)

Пример фрагмента расширенного отчета:
![image](https://user-images.githubusercontent.com/44323790/142164528-a3a3eca0-0af1-459e-831a-e12505ad7324.png)
