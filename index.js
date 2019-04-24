"use strict";

// импорт библиотек
const fs = require('fs');
const xlsx = require('node-xlsx').default;

////////////////////////////////////////////////////

// Получение содержимого xlsx файла

// имя файла
const fileName = "File.xlsx";
// получаем содержимое файла
const fileContent = fs.readFileSync(fileName);

// содержимое в виде объекта
const obj = xlsx.parse(fileContent)[0];

// массив строк
const rowsArray = obj.data;

// пробегаемся по всем строкам
for(let i = 0; i < rowsArray.length; i++) {
    // получаем строку
    const row = rowsArray[i];
    // выводим содержимое строки
    console.log(row[0] + " " + row[1] + " " + row[2]);
}

////////////////////////////////////////////////////

// Создание собственного xlsx файла

// содержимое таблицы в виде массива
// каждая строка таблицы является массивом
const content = [
    ["Фамилия", "Имя", "Оценка"],
    ["Иванов", "Иван", 4],
    ["Петров", "Пётр", 5],
    ["Сидоров", "Семён", 4],
    ["Орлов", "Олег", 3],
    ["Дмитриев", "Дима", 5],
];

// задаём ширину каждой колонки
const option = {
    '!cols': [
        { wch: 30 },
        { wch: 30 }, 
        { wch: 20 },
    ]
};

// имя
const name = "table";
// содержимое
const data = content;

// формируем объект
const objNameData = {
    name: name,
    data: data,
}

// получаем содержимое таблицы
const buffer = xlsx.build([
    objNameData
], option); 

// сохраняем содержимое в файл
fs.writeFileSync("My.xlsx", buffer);

////////////////////////////////////////////////////

