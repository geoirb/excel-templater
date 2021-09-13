# TEMPLATER

Сервис принимает из очереди сообщений json документ и по нему заполняет необходимый шаблон и посылает результат в очередь сообщений.

Для запуска используются следующие переменные окружения:

MQ_HOST - хост, на котором расположена кафка, по дефолту localhost
MQ_PORT - порт для подключения к кафке
VALUE_ARE_REQUIRED - флаг определяющий то, что все значения в шаблоне обязательные
TEMPLATE_DIR - папка, в которой находятся шаблоны
FILL_IN_TOPIC_REQUEST - топик кафки с запросами на заполнение шаблона
FILL_IN_TOPIC_RESPONSE - топик кафки с результатами заполнения


Структура входного json докумета
```json
{
  "user_id": int,
  "uuid": string,
  "template": string,
  "payload": object
}
```
> uuid - идентификатор запроса
>
> user_id идентификатор пользователя отправившего запрос
> 
> template - имя шаблона, который необходимо заполнить
>
> payload - объект с данными, которыми необходимо заполнить шаблон

Структура результата
```json
{
    "is_success": bool,
    "user_id": int,
    "uuid": string,
    "document": binary,
    "error": string
}
```

> is_success - признак успешности заполнения
> 
> user_id идентификатор пользователя отправившего запрос
> 
> uuid - идентификатор запроса
> 
> document - заполненный документ, передается если is_success = true
>
> payload - текст ошибки, передается если is_ok = false

## Placeholders

Для заполнения документа необходимо в нем расставить плейсхолдеры в местах заполнения, по которым программа найдет соответствующие данные в json файле

Плейсхолдер - это специально построенная строка, с помощью которой шаблонизатор понимает куда и что ему необходимо вставить

1. плейсхолдер состоит из ключей, разделенных знаком `:`
2. ключи могут быть нескольких видов:
   * field name
   * array 
   * qr code 
   * image
3. плейсхолдер находится между фигурными скобками {}

Ключи:

1. `field name` - строка типа [a-zA-Z0-9_]+,ключ необходимый для спуска по дереву json, представляет собой имя полей json, по последнему ключу должно лежать значение
2. `array` - служит для указания генерации таблицы из массива, **данный ключ находится в конце плейсхолдера**, плейсхолдер располагается в левом верхнем углу генерируемой таблицы, на следующей строке расположена шапка таблицы, элементы шапки это плейсхолдеры для получения данных из элементов массива, они указывают какие элементы куда всталять, показывают размеры таблицы и задают стили элементов таблицы. Начало шапки таблицы располагается под соответсвующим плейсхолдером с ключом array
**!!!строка с плейсхолдером `array` и строка с шапкой таблицы буду удалены!!!**
3. `qr code` - строка типа qr_code_[0-9]+, указывает, что нужно вставить массив qr кодов начиная с позиции ключа,  **данный ключ находится в конце плейсхолдера**. Данные берутся из json файла, название поля должно соответствовать ключу. Данные должны быть в представлены как массив строк.
4. `image` - строка типа image_[0-9]+, указывает, что нужно вставить `.png` изображение на данное место, данные для изображения лежат в json по плейсхолдеру

Если программа не находит данные в json по плейсхолдеру и `VALUE_ARE_REQUIRED`, при запуске программы, выставлен в `false`, то на месте плейсхолдер заменяется на пустое значение, если `VALUE_ARE_REQUIRED`, при запуске программы, выставлен в `true`, программа возращает ошибку.

Пример:

`payload` запроса:
```json
{
  "field_name_0": "value_field_name_0",
  "field_name_1":{
    "field_name_3": "value_field_name_3",
    "data":[
      {
        "1":1.0,
        "2_0":2.0,
        "3_0":3.0
      },
      {
        "1":1.1,
        "2_0":2.1,
        "3_0":3.1
      },
      {
        "1":1.2,
        "2_0":2.2,
        "3_0":3.2
      }
    ],
    "qr_code_0":[
      "qr-code-0-0",
      "qr-code-0-1"
    ]
  }
}
```

шаблон

![шаблон](images/template.png)

результат

![результат](images/result.png)