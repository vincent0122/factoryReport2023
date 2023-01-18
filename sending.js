function sending() {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var setReport = ss.getSheetByName("일지작성");
    var stock = ss.getSheetByName("재고");
    var stockLastRow = stock.getLastRow();
    var stockChange = ss.getSheetByName("재고변동");
    var etc = ss.getSheetByName("기타활동");
    var date = setReport.getRange("b1").getValue();

    const setLastRow = (sheetName) => {
        var emp = sheetName.getRange("b1:b3000").getValues();
        for (var lastRow = 0; lastRow < emp.length; lastRow++) {
            if (!emp[lastRow].join("")) {
                return lastRow;
            }
        }
    };
    var stockChangeLastRow = setLastRow(stockChange);
    var etcLastRow = setLastRow(etc);

    const todayStockSolid = (sheetName) => {};

    const nameCleaning = (str) => {
        var myArray = str.split(/[\s,]+/);
        return [myArray.join(",")];
    };

    const Values = {
        muguri: setReport.getRange("K3:P3").getValues(),
        produce: setReport.getRange("A4:S28").getValues(),
        etcAct: setReport.getRange("A31:K34").getValues(),
        ilyong: setReport.getRange("N31:Q33").getValues(),
        came: setReport.getRange("A37:F40").getValues(),
        jidaekpp: setReport.getRange("I37:K40").getValues(),
        vacation: setReport.getRange("a44:a51").getValues(),
        vacationHalf: setReport.getRange("b44:b51").getValues(),
        cost: setReport.getRange("d44:e51").getValues(),
        back: setReport.getRange("I44:K51").getValues(),
        out: setReport.getRange("N37:R56").getValues(),
        stock: stock.getRange(1, 1, stockLastRow, 5).getValues(),
    };

    const addMinus = (item, qty, options) => {
        var stockData = Values.stock;
        var indexOfItem = stockData.findIndex((subArray) =>
            subArray.includes(item)
        );
        Logger.log(indexOfItem);

        if (
            options === "생산" ||
            options === "반품" ||
            options === "구매" ||
            options === "재고증가" ||
            options === "수입" ||
            options === "지대입고"
        ) {
            stock
                .getRange(indexOfItem + 1, 4)
                .setValue(stock.getRange(indexOfItem + 1, 5).getValue());
            stock
                .getRange(indexOfItem + 1, 5)
                .setValue(stock.getRange(indexOfItem + 1, 5).getValue() + qty);
        } else if (
            options === "사용" ||
            options === "출고" ||
            options === "손실" ||
            options === "재고감소" ||
            options === "지대사용"
        ) {
            stock
                .getRange(indexOfItem + 1, 4)
                .setValue(stock.getRange(indexOfItem + 1, 5).getValue());
            stock
                .getRange(indexOfItem + 1, 5)
                .setValue(stock.getRange(indexOfItem + 1, 5).getValue() - qty);
        }

        var todayStock = stock.getRange(indexOfItem + 1, 5).getValue();
        return todayStock;
    };

    const cleaningProduct = () => {
        var produce = Values.produce;
        if (produce[2][0] != "") {
            var indexRemove = [0, 1, 4, 5, 8, 9, 13, 14, 17, 18, 21, 22];
            for (var i = indexRemove.length - 1; i >= 0; i--) {
                var indexToRemove = indexRemove[i];
                produce.splice(indexToRemove, 1);
            }
            var produce = produce.filter((el) => el[2] != "");
            for (var i = 0; i < produce.length; i++) {
                if (produce[i][0] === "") {
                    ui.alert(`${i + 1}번째 제품명이 비어 있습니다`);
                    return;
                }

                produce[i][16] = nameCleaning(produce[i][16]);

                for (var j = 0; j < 12; j++) {
                    if (produce[i][j] === "" || j === 11) {
                        produce[i].push(j);
                        break;
                    }
                }

                var inputData = [
                    [
                        date,
                        "생산",
                        produce[i][0],
                        produce[i][1], ,
                        produce[i][16],
                        produce[i][14],
                        produce[i][15], , , ,
                        produce[i][17],
                        produce[i][18],
                    ],
                ];
                var todayStock = addMinus(
                    inputData[0][2],
                    inputData[0][3],
                    inputData[0][1]
                );
                inputData[0][4] = todayStock;

                var inputData2 = [
                    [date, "지대사용", produce[i][13], , , produce[i][12]],
                ];
                var todayStock2 = addMinus(
                    inputData2[0][5],
                    inputData2[0][2],
                    inputData2[0][1]
                );
                inputData2[0][4] = todayStock2;

                stockChange
                    .getRange(stockChangeLastRow + 1, 1, 1, 13)
                    .setValues(inputData);
                stockChangeLastRow = stockChangeLastRow + 1;

                etc.getRange(etcLastRow + 1, 1, 1, 6).setValues(inputData2);
                etcLastRow = etcLastRow + 1;

                for (var k = 2; k < produce[i][19]; k = k + 2) {
                    var inputData = [
                        [
                            date,
                            "사용",
                            produce[i][k],
                            produce[i][k + 1], , , , ,
                            produce[i][0],
                        ],
                    ];
                    var todayStock = addMinus(
                        inputData[0][2],
                        inputData[0][3],
                        inputData[0][1]
                    );
                    inputData[0][4] = todayStock;

                    stockChange
                        .getRange(stockChangeLastRow + 1, 1, 1, 9)
                        .setValues(inputData);
                    stockChangeLastRow = stockChangeLastRow + 1;
                }
            }
        }
        return produce;
    };

    const cleaningOut = () => {
        var out = Values.out;
        var out = out.filter((el) => el[0] != "");
        Logger.log(out[0][0]);
        Logger.log(out[0][1]);
        Logger.log(out[1][1]);
        //[제품명, 수량, 비고(거래처명)]

        for (var i = 0; i < out.length; i++) {
            var inputData = [
                [date, "출고", out[i][0], out[i][3], , , , , , , , , out[i][4]],
            ];
            var todayStock = addMinus(
                inputData[0][2],
                inputData[0][3],
                inputData[0][1]
            );
            inputData[0][4] = todayStock;

            stockChange
                .getRange(stockChangeLastRow + 1, 1, 1, 13)
                .setValues(inputData);
            stockChangeLastRow = stockChangeLastRow + 1;
        }
    };

    cleaningProduct();
    cleaningOut();
}