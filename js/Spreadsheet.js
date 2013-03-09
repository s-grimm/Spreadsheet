﻿(function ($) {
    $.fn.Spreadsheet = function (options) {
        //settings for spreadsheet app
        var settings = $.extend(
            {
                'columns': 10,
                'rows': 20,
            }, options);

        function labelCode(i) { //make i 0 based (ie 0 = A, 1 = B, 25 = Z, 26 = AA)
            var r = parseInt((i / 26), 10);
            if (r == 0) {
                return String.fromCharCode(65 + i);
            }
            else {
                return labelCode(r - 1) + '' + String.fromCharCode(65 + (i % 26));
            }
        }

        function decodeLabel(ch) {
            if (ch.length == 1) {
                return ch.charCodeAt(0) - 64; //offset to 64 because index of columns does not start at 0
            } else {
                return ((ch.charCodeAt(0) - 64) * 26) + decodeLabel(ch.substr(1));
            }
        }

        function updateSelectedCell(cellRef) {
            var $formulabar = $('.formula-bar-text');
            cellRef.val($formulabar.val());
        }

        function updateFormulaBar(text) {
            var $formulabar = $('.formula-bar-text');
            $formulabar.val(text);
        }

        function getCellValueFromCoordinates(xPos, yPos) {
            var row = $('.spreadsheet tbody tr').eq(yPos);
            var cell = row.children('td').eq(decodeLabel(xPos) - 1);
            var inp = $(cell[0]).children("input[type='text']:first").val();

            return inp;
        }
        function updateLeSpreadsheet(cellRef) {//string of cell (ie A1)
            cellRef = cellRef.trim(); //remove any whitespace before and after this
            var items = cellRef.match(/(\D+)(\d+)/).slice(1);
            var val = getCellValueFromCoordinates(items[0], items[1]);
        }

        function recalculateAllCells() {
            $('.spreadsheet-input-cell').each(function () {
                var $this = $(this);
                if ($this.data('formula') != "" && $this.data('formula') != undefined && $this.data('formula') != 'undefined') {
                    var oldVal = $this.val();
                    calculateCellValue($this);
                    var newVal = $this.val();
                    if (oldVal != newVal) {
                        recalculateAllCells();
                    }
                }
            });
        }

        function calculateCellValue(inputElement) {
            var val = inputElement.val();
            if (inputElement.data('formula') != "" && inputElement.data('formula') != undefined && inputElement.data('formula') != 'undefined') {
                val = inputElement.data('formula');
                if (val.match(/\=[Ss][Uu][Mm]/)) {
                    //we are in sum formula, split on : after removing =SUM, '(' and ')'
                    val = val.replace(/\s+/, '');
                    val = val.replace(/\=[Ss][Uu][Mm]/, '');
                    val = val.replace(/\)/, '');
                    val = val.replace(/\(/, '');
                    val = val.trim();
                    values = val.split(':');

                    for (var i = 0; i < values.length; ++i) {
                        values[i] = values[i].trim(); //remove any whitespace before and after this
                        values[i] = values[i].match(/(\D+)(\d+)/).slice(1);
                    }
                    var sum = 0;
                    //calculate highest and lowest cell ranges
                    var lowerCol = decodeLabel(values[0][0]) < decodeLabel(values[1][0]) ? decodeLabel(values[0][0]) : decodeLabel(values[1][0]);
                    var higherCol = decodeLabel(values[0][0]) > decodeLabel(values[1][0]) ? decodeLabel(values[0][0]) : decodeLabel(values[1][0]);
                    var lowerRow = values[0][1] < values[1][1] ? values[0][1] : values[1][1];
                    var higherRow = values[0][1] > values[1][1] ? values[0][1] : values[1][1]
                    //sum them
                    for (var i = lowerRow; i <= higherRow; ++i) {
                        for (var x = lowerCol; x <= higherCol; ++x) {
                            sum += parseInt(getCellValueFromCoordinates(labelCode(x - 1), i));
                        }
                    }
                    inputElement.val(sum);
                } else {
                    //we are in a straight = formula. split on operators (/[\+\/\-\*]/)
                }
            }
        }

        function cellLostFocus(inputElement) {
            var val = inputElement.val();
            if (val == "" || val == 'undefined' || val === undefined) {
                if (inputElement.data('formula') != "" && inputElement.data('formula') != undefined && inputElement.data('formula') != 'undefined') {
                    inputElement.data('formula', ''); //if the cell content is empty, reset the formula if there is one (aka you deleted the cell contents)
                }
                return; //dont go beyond this if the cell is empty
            }
            if (val[0] == '=') {
                inputElement.data('formula', val.toUpperCase());
            }
            if (inputElement.data('formula') != "" && inputElement.data('formula') != undefined && inputElement.data('formula') != 'undefined') {
                calculateCellValue(inputElement);
            }
            //calculation of this cells formula is complete, update the rest of the table
            recalculateAllCells();
        }

        function clearAllCells() {
            $('.spreadsheet-input-cell').each(function () {
                var $this = $(this);
                if ($this.data('formula') != "" && $this.data('formula') != undefined && $this.data('formula') != 'undefined') {
                    $this.data('formula', '');
                }
                $this.val('');
            });
        }

        //rename this...
        function postShit() {
            var jsonCells = "{";
            var counter = 0;
            var col = 0;

            var isFirst = true;
            // whatToStringify[3][counter][col]
            $('.spreadsheet-table tr').each(function () {
                var $this = $(this);
                
                $this.children('input').each(function () {
                    if(!isFirst){
                        jsonCells += ","
                    }else{
                        isFirst = false;
                    }
                    var $this = $(this);
                    var cellPrefix = labelCode(col);
                    if ($this.data('formula') != "" && $this.data('formula') != undefined && $this.data('formula') != 'undefined') {
                        jsonCells += "'"+cellPrefix+counter+"':'" + $this.data('formula') + "'";
                    } else {
                        jsonCells += "'" + cellPrefix + counter + "':'" + $this.val() + "'";
                    }
                    col++;
                });
                col = 0;
                counter++;
            });
            jsonCells += "}";
            var saveValue = "{ 'rows':'" + settings.rows + "', 'columns':'" + settings.columns + "' , cells:" + jsonCells + " }";

            $.ajax({
                type: 'POST',
                url: 'SaveSpreadsheet.asmx/HelloWorld',
                data: "{'name':'" + saveValue + "'}",
                contentType: 'application/json; charset=utf-8',
                dataType: 'json',
                success: function () {
                    alert('WOOT');
                },
                error: function (xhr, ajaxOptions, thrownError) {
                    alert('Error : ' + xhr.status + ' - ' + xhr.statusText + ' : ' + thrownError);
                }
            });
        }

        return this.each(function () {
            //reference scope to local variable so you don't have to do scope it each time.
            var $this = $(this);
            $this.addClass('spreadsheet');
            //icon-fullscreen
            var $formulaBar = $('<div class="input-prepend input-append span12"><div class="btn-group"><button class="btn dropdown-toggle" data-toggle="dropdown"><i class="icon-list-alt"></i><span class="caret"></span></button><ul class="dropdown-menu"><li><a class="formula-sum">Sum</a></li></ul></div><input type="text" class="span6 formula-bar-text"/><div class="btn-group"><button class="btn dropdown-toggle" data-toggle="dropdown"><i class="icon-edit"></i><span class="caret"></span></button><ul class="dropdown-menu"><li><a class="btn-clear-spreadsheet">Clear</a></li><li><a class="btn-load-spreadsheet">Load</a></li><li><a class="btn-save-spreadsheet">Save</a></li></ul></div></div>');
            $this.append($formulaBar);
            $this.append('<br />'); //fix for styling in firefox

            var $table = $('<table class="spreadsheet-table">');
            $this.append($table);

            var $rowHeader = $('<tr><th class="corner-cell"></th></tr>');
            $table.append($rowHeader);
            for (var i = 0; i < settings.columns; ++i) {
                //String.fromCharCode
                var $th = $('<th>' + labelCode(i) + '</th>');
                $rowHeader.append($th);
            }

            for (var rCount = 1; rCount <= settings.rows; rCount++) {
                var $row = $('<tr>');
                $table.append($row);
                $row.append('<th>' + rCount + '</th>');
                for (var cCount = 1; cCount <= settings.columns; cCount++) {
                    var $col = $('<td>');
                    $row.append($col);
                    $col.append('<input type="text" class="spreadsheet-input-cell"/>');
                }
            }

            $('.formula-sum').on('click', function () {
                var leCell = $('.selected');
                leCell.val('=SUM(' + leCell.val() + ')');
                updateFormulaBar(leCell.val());
                leCell.data('formula', leCell.val());
            });

            $('.btn-clear-spreadsheet').on('click', function () {
                clearAllCells();
            });

            $('.btn-save-spreadsheet').on('click', function () {
                postShit();
            });

            $('.formula-bar-text').on('focusout', function () {
                var leCell = $('.selected');
                calculateCellValue(leCell);
            });

            $('.spreadsheet-input-cell').on('click', function () {
                $('.selected').removeClass('selected').css('background-color', 'white');
                var $selected = $(this);
                $selected.closest('input').addClass('selected').css('background-color', 'lightblue');
                var leCell = $('.selected');
                if (leCell.data('formula') == "" || leCell.data('formula') == undefined || leCell.data('formula') == 'undefined') {
                    updateFormulaBar(leCell.val());
                } else {
                    updateFormulaBar(leCell.data('formula'));
                    updateSelectedCell(leCell);
                }
                var $td = $selected.closest('td'),
                rowIndex = $td.closest('tr').index(),
                colIndex = $td.index(); //this is not 0 based : this is a +1 based input (A = 1, Z=26, etc)
                //bind events to selected input and formula bar
                $formulaBar.unbind('keyup');
                $('.spreadsheet-input-cell').unbind('keyup');
                $formulaBar.keyup(function (e) {
                    var code = (e.keyCode ? e.keyCode : e.which);
                    if (code == 13) { //Enter keycode
                        //Do something
                    } else {
                        updateSelectedCell(leCell);
                    }
                });

                leCell.keyup(function (e) {
                    var code = (e.keyCode ? e.keyCode : e.which);
                    if (code == 13) { //Enter keycode
                        //Do something
                        updateLeSpreadsheet(labelCode(colIndex - 1), rowIndex, leCell);
                    } else {
                        updateFormulaBar(leCell.val());
                    }
                });
            }).on('focusout', function () {
                var leCell = $('.selected');
                cellLostFocus(leCell);
            });
        });
    };
})(jQuery);