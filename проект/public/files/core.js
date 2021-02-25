//-------------ADD AND REMOVE FIELDS

class sideBarState {
    constructor() {
        this.fields = document.getElementsByClassName("sidebar-fields");
        this.length = this.fields.length;
        this.sidebar = document.getElementsByClassName("sidebar")[0];
    }

    refresh() {
        this.fields = document.getElementsByClassName("sidebar-fields");
        this.length = this.fields.length;
        let numBlocks = document.getElementsByClassName("numeric")
        if (this.length) {
            for (let i = 0; i < this.length; i++) {
                numBlocks[i].textContent = "№" + (i + 1);
            }
        }
    }
}

let addElmBtn = document.getElementsByClassName("sidebar-text-add")[0];
let state = new sideBarState();
addElmBtn.onclick = () => addFieldCopy(state)

function removeField(state, elm) {
    elm.parentNode.remove()
    state.refresh();
}

function addFieldCopy(state) {
    let sideBar = state.sidebar;
    sideBar.insertAdjacentHTML("beforeend", `
                <div class="sidebar-fields">
                <span class="numeric"></span><span class="sidebar-fields-inputcontainer">
                    <input type="text" class="sidebar-fields-ordernumber">
                </span>
                <span class="sidebar-fields-remove" onselectstart="return false" onclick="removeField(state, this)">Удалить</span>
                <span class="sidebar-fields-inputcontainer">
                <input type="text" class="sidebar-fields-ranges">
                </span>
                </div>`)
    state.refresh();
}


// EXCEL PART

class ExcelBooks {
    constructor(pointsPath, protoPath) {
        this.pointsPath = pointsPath;
        this.constPath = this.findConstList();
        this.protoPath = protoPath;
        return (async () => {
            this.pointsBook = await this.getBook(this.pointsPath);
            this.constBook = await this.getBook(this.constPath);
            this.protoBook = await this.getBook(this.protoPath);
            return this
        })()
    }

    findConstList() {
        let files = fs.readdirSync("./");
        for (let file of files) {
            if (/[.]xlsx?/.exec(file) && file !== "template.xlsx") return "./" + file
        }
    }

    async getBook(path) {
        try {
            let constBook = new Excel.Workbook();
            await constBook.xlsx.readFile(path);
            return constBook
        } catch (err) {
            console.log(err)
            alert("Не удалось найти файл с точками либо файл с конструкциями либо файл шаблона. " +
                "Необходимо закинуть файл с конструкциями и файл шаблона протокола в папку с приложением, " +
                "путь к файлу с точками указывается в интерфейсе приложения.")
        }
    }
}

class Protocol {
    constructor(number, range, testDate, whoTested, whoControlled) {
        this.number = number;
        this.range = range.split("-");
        this.testDate = testDate;
        this.whoTested = whoTested;
        this.whoControlled = whoControlled;
        this.objectName = "";
        this.concretingDate = "";
        this.concreteMark = "";
        this.constInfo = "";
        this.points = "";
        this.isEmpty = false;
    }

    getInfoData(workbook) {
        if (!workbook.worksheets.length) {
            alert("Вероятно, файл реестра имеет расширение отличное от xlsx")
            throw new Error("Отсутствуют данные в книге с реестром")
        }
        let worksheet = workbook.getWorksheet("физ-мех хар. бетона 21г");

        function getInfoFromRow(row) {
            if (resultRow === null) return null
            let fileInfo = Object.assign({}, columns)
            for (let col of Object.keys(columns)) {
                let cell;
                if (col !== "concretingDate") {
                    cell = worksheet.getCell(row, columns[col]).value
                } else {
                    try {
                        cell = worksheet.getCell(row, columns[col]).value.toLocaleDateString()
                    } catch (err) {
                        let value = worksheet.getCell(row, columns[col]).value
                        let longYeardateMatch = value.match(/\d+[.]\d+[.]\d{4}/i)
                        let shortYeardateMatch = value.match(/\d+[.]\d+[.]\d{2}/i)
                        let dateMatch = longYeardateMatch || shortYeardateMatch;
                        if (dateMatch) {
                            cell = dateMatch === shortYeardateMatch ? dateMatch[0].replace(/(\d+[.]\d+[.])(\d{2})/i, "$120$2") : dateMatch[0]
                        } else {
                            alert("Неверный формат даты в реестре конструкций. Вместо даты бетонирования будет поставлен мой день рождения")
                            cell = "10.08.1995"
                        }
                    }
                }
                fileInfo[col] = cell;
            }
            fileInfo.concreteMark = fileInfo.concreteMark.match(/(B|В|M|М)\d+([.,]\d+)?/i)[0];
            return fileInfo
        }

        let columns = {
            number: 2,
            objectName: 5,
            concreteMark: 8,
            concretingDate: 9,
            constInfo: 19
        }
        let startRow = 20;
        let valueArray = [];
        let resultRow = null;
        for (let i = startRow; ; i++) {
            let value = worksheet.getCell(i, columns.number).value
            if (value == this.number) {
                resultRow = i;
                break
            }
            valueArray.push(value);
            if (valueArray.filter(elm => elm === null).length > 100) {
                this.isEmpty = true;
                break
            }
        }
        return Object.assign(this, getInfoFromRow(resultRow))
    }

    getPointsData(workbook) {
        if (!workbook.worksheets.length) {
            alert("Вероятно, файл с точками имеет расширение отличное от xlsx")
            throw new Error("Отсутствуют данные в книге с точками")
        }
        let worksheet = workbook.getWorksheet(1);
        let startRow = 2;
        let column = 9;
        let points = getPoints(column)

        function getPoints(col, row = startRow) {
            let points = [];
            let point;
            while (point = worksheet.getCell(row, col).value) {
                row++;
                points.push(point)
            }
            return points
        }

        function replacePoints(points) {
            function numbersIntoPercents(array) {
                let average = array.reduce((a, b) => a + b) / array.length
                array = array.map(elm => Math.abs(100 - (elm * 100 / average)))
                return array
            }

            let pos = 20
            if (points.length < pos) return points
            let checked = points.slice(0, pos)
            let average = checked.reduce((a, b) => a + b) / checked.length
            let percentArray = numbersIntoPercents(checked)
            while (percentArray.some(elm => elm > 10)) {
                if (points[pos]) {
                    let max = Math.max(...percentArray);
                    let index = percentArray.indexOf(max);
                    checked.splice(index, 1)
                    checked.push(points[pos])
                    average = checked.reduce((a, b) => a + b) / percentArray.length
                    pos++;
                    percentArray = numbersIntoPercents(checked)
                } else {
                    alert(`В протоколе ${this.number} не удалось автоматически найти 20 подходящих точек, соответствующих условию отклонения 10% от среднего значения. Точки 
                       этого протокола будут проставлены по порядку, начиная с первой.`)
                    return points
                }
            }
            return checked;
        }

        let rangePoints = points.slice(this.range[0] - 1, this.range[1])
        let replacedPoints = replacePoints.call(this, rangePoints);

        this.points = replacedPoints;
        if (this.points.length < 20) {
            let agreed = confirm(`В протоколе №${this.number} обнаружено менее 20 точек. Проверьте правильность заполнения полей с точками.` +
                `Нажимая ОК программа продолжит свое выполнение, в противном случае программа прекратит запись и ни один протокол не будет записан.`)
            if (!agreed) throw new Error("Программа завершена по инициативе пользователя")
        }
        return replacedPoints
    }

    async write(workbook, path) {
        if (!workbook.worksheets.length) {
            alert("Вероятно, шаблон протокола имеет расширение отличное от xlsx")
            throw new Error("Отсутствуют данные в книге с протоколом")
        }
        let worksheet = workbook.getWorksheet("Протокол");
        let cells = {
            whoTested: [71, 14],
            whoControlled: [72, 14],
            number: [80, 2],
            concreteMark: [12, 44],
            points: [79, 9],
            constInfo: [96, 24],
            concretingDate: {
                Day: [38, 30],
                Month: [38, 32],
                Year: [38, 34]
            },
            testDate: {
                Day: [66, 30],
                Month: [66, 32],
                Year: [66, 34]
            }
        }
        for (let cell of Object.keys(cells)) {
            if (Array.isArray(cells[cell])) {
                if (cell === "points") {
                    pastePoints.call(this, cells[cell])
                    continue
                }
                worksheet.getCell(...cells[cell]).value = this[cell];
            } else {
                let dateObj = cells[cell]
                let arrayFromDate = this[cell].split(".")
                let keys = Object.keys(dateObj);
                for (let i = 0; i < arrayFromDate.length; i++) {
                    worksheet.getCell(...dateObj[keys[i]]).value = arrayFromDate[i];
                }
            }
        }

        function pastePoints(cell) {
            for (let i = 0; i < this.points.length; i++) {
                if (i > 19) break
                worksheet.getCell(cell[0] + i, cell[1]).value = this.points[i]
            }
        }

        function getDifferenceFromDatesInDays(date1, date2) {
            date1 = date1.split(".").reverse()
            date1[1] = Number(date1[1]) - 1
            date2 = date2.split(".").reverse()
            date2[1] = Number(date2[1]) - 1
            let dateDifference = new Date(...date2) - new Date(...date1)
            let days = dateDifference / (1000 * 60 * 60 * 24)
            return days
        }

        function findNameOfBuilding(info) {
            let found;
            if (info.match(/контурн[^\s]* стен/)) found = "контурные стены"
            else if (info.match(/внутренн[^\s]* стен/)) found = "внутренние стены"
            else if (info.match(/плит/)) found = "плита"
            else if (info.match(/фундамент/)) found = "фундамент"
            else if (info.match(/колонн/)) found = "колонны"
            else if (info.match(/стен/)) found = "стена"
            else if (info.match(/секци/)) found = "секция"
            else found = "КОНСТРУКЦИЯ"
            return found
        }

        if (!this.concreteMark) alert(`не удалось корректно преобразовать марку бетона в протоколе с номером ${this.number}`)
        let difference = getDifferenceFromDatesInDays(this.concretingDate, this.testDate)
        let name = this.constInfo.match(/\d{2}\w{3}/)[0];
        let nameOfBuilding = findNameOfBuilding(this.constInfo)
        let fileName = `${name} №${this.number}-${difference} ${nameOfBuilding} от ${this.concretingDate}.xlsx`
        await workbook.xlsx.writeFile(path + fileName)
    }
}

class ProtocolGenerator {
    constructor() {
        this.whoTested = document.getElementById("nikimt").value;
        this.whoControlled = document.getElementById("ase").value;
        this.testDate = document.getElementsByClassName("content-date-field")[0].value.split("-").reverse().join(".");
        this.ranges = Array.from(document.getElementsByClassName("sidebar-fields-ranges")).map(elm => elm.value);
        this.numbersOfConst = document.getElementsByClassName("sidebar-fields-ordernumber");
        this.protoPath = "./template.xlsx"
        this.pointsPath = document.getElementsByClassName("content-file-path")[0].textContent;
        this.writePath = document.getElementsByClassName("content-file-path")[1].textContent;
        this.needChangeSamePoints = document.getElementsByClassName("checkbox")[0].checked
        this.needToChangePoints = document.getElementsByClassName("checkbox")[1].checked
        this.emptyProtos = [];
    }

    checkFields() {
        function anyProblemWithFields(num) {
            if (!this.numbersOfConst[num].value && !Number(this.numbersOfConst[num].value)) return `Номер заявки должен быть числом в поле №${num + 1}`
            if (!/^\d+[-]\d+$/.test(this.ranges[num])) return `Неверный формат записи диапазона в поле №${num + 1}`
            return false
        }

        function anyProblemWithStaticElms() {
            if (!this.testDate) return "Необходимо указать дату испытания"
            if (!this.whoTested) return "Необходимо указать сотрудника НИКИМТ"
            if (!this.whoControlled) return "Необходимо указать сотрудника АСЭ"
            return false
        }

        let error = anyProblemWithStaticElms.call(this);
        if (error) {
            alert(error);
            return false;
        }
        for (let i = 0; i < this.ranges.length; i++) {
            error = anyProblemWithFields.call(this, i);
            if (error) {
                alert(error);
                return false
            }
        }
        return true
    }

    getProtocols(books) {
        let length = this.ranges.length
        let arrayWithProto = new Array(length);
        for (let i = 0; i < length; i++) {
            let num = this.numbersOfConst[i].value;
            let proto = new Protocol(num, this.ranges[i], this.testDate, this.whoTested, this.whoControlled);
            console.log(proto)
            proto.getInfoData(books.constBook)
            proto.getPointsData(books.pointsBook)
            if (proto.isEmpty) this.emptyProtos.push(num)
            arrayWithProto[i] = proto
        }
        return arrayWithProto
    }

    async fixPointsBook(books) {
        if (!books.pointsBook.worksheets.length) {
            alert("Вероятно, файл с точками имеет расширение отличное от xlsx")
            throw new Error("Отсутствуют данные в книге с точками")
        }
        let worksheet = books.pointsBook.getWorksheet(1);
        let startRow = 2;
        let columnsToChange = {
            K: 8,
            material: 7
        }
        let pointsColumn = 9;
        let pointsCells = getPointsCells(pointsColumn);
        if (this.needToChangePoints) {
            for (let range of this.ranges) {
                let [from, to] = range.split("-").map(elm => Number(elm));
                from = from - 1
                let partOfPointsCells = pointsCells.slice(from, to)
                partOfPointsCells = removeHigherAndLower(partOfPointsCells);
                for (let i = from; i < partOfPointsCells.length; i++) {
                    pointsCells[i] = partOfPointsCells[i];
                }
            }
        }
        try {
            fillCellsWithFixed.call(this, pointsCells);
            await books.pointsBook.xlsx.writeFile(this.pointsPath)
        } catch (err) {
            alert("Возникла проблема в файле с точками. Вероятно, где-то присутствует пустое значение, либо указан неверный интервал.")
            throw err
        }

        function removeHigherAndLower(pointsCells) {
            let topLimit = 5000;
            let botLimit = 3900;
            pointsCells = pointsCells.map(elm => {
                elm.value = (elm.value < topLimit && elm.value > botLimit) ? elm.value : null
                return elm
            })
            let filteredPointsCells = pointsCells.filter(elm => elm.value !== null);
            let average = Math.floor(filteredPointsCells.reduce((a, b) => a + b.value, 0) / filteredPointsCells.length)
            pointsCells = pointsCells.map(elm => {
                let delta = -75 + Math.floor(Math.random() * 150);
                let corrected = (average + delta > topLimit) || (average + delta < botLimit) ? average: average + delta
                    elm.value = elm.value === null ? corrected : elm.value
                return elm
            })
            return pointsCells
        }

        function getPointsCells(col, row = startRow) {
            let cells = [];
            while (worksheet.getCell(row, col).value) {
                cells.push(worksheet.getCell(row, col))
                row++;
            }
            return cells
        }

        function fillCellsWithFixed(pointsCells) {
            for (let col of Object.keys(columnsToChange)) {
                for (let i = startRow; i < startRow + pointsCells.length; i++) {
                    if (col === "K") worksheet.getCell(i, columnsToChange[col]).value = 1;
                    else worksheet.getCell(i, columnsToChange[col]).value = "Бетон тяжелый";
                }
            }

            for (let range of this.ranges) {
                let [from, to] = range.split("-").map(elm => Number(elm));
                let arrayFromRange = pointsCells.slice(from - 1, to).map(elm => elm.value)
                let changedArray = changeTheSame(arrayFromRange);
                for (let i = from - 1; i < to; i++) {
                    pointsCells[i].value = changedArray[i - (from - 1)]
                }
            }
        }


        function changeTheSame(arr) {
            let changed = true;
            while (changed) {
                changed = false;
                for (let i = 0; i < arr.length; i++) {
                    for (let j = 0; j < arr.length; j++) {
                        if (arr[i] === arr[j] && i !== j) {
                            arr[j]++
                            changed = true
                        }
                    }
                }
            }
            return arr
        }
    }

    async run() {
        let books = await new ExcelBooks(this.pointsPath, this.protoPath);
        if (this.needChangeSamePoints) await this.fixPointsBook(books)
        let protocols = this.getProtocols(books)
        await protocols.filter(elm => !elm.isEmpty).forEach(elm => {
            elm.write(books.protoBook, this.writePath)
        })
    }

}


const Excel = require('exceljs');
let fs = require("fs");
let startBtn = document.getElementsByClassName("content-crate-button")[0];

startBtn.onclick = async () => {
    let app = new ProtocolGenerator();
    if (!app.checkFields()) {
        return
    }
    await app.run();
    if (app.emptyProtos.length) {
        alert(`Протоколы под номерами:${app.emptyProtos.join(",")} не были созданы, так как они отсутствуют в реестре конструкций`)
    }
}


