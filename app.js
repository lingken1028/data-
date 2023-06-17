$(document).ready(function() {
  var data = []; // 存储Excel数据的数组
  var displayData = []; // 显示的数据
  var workbook; // 存储Excel工作簿对象

  
  // 选择Excel文件并读取数据
  $('#fileInput').click(function() {
   
    $.ajax({
      url: 'https://script.googleusercontent.com/macros/echo?user_content_key=pGsvIjTEMnY1mQvWdB686ETL1kH1sAav9CGDhdk6Vav_AK6O_AjodNsTsSaSJR5NlND-20XOviQTvdqJpqyeWFdCXzYJsboUm5_BxDlH2jW0nuo2oDemN9CCS2h10ox_1xSncGQajx_ryfhECjZEnL1AgOGVUnVifADHJtysPXa6btAjAVZ4dMFguehaPME3aQHf0zi23KMRM_pjlOqlEyuH8d1oT6Hk29bd0_3O8KQInECzU10futz9Jw9Md8uu&lib=MyNC0aJR-Tl-mJWJQ_5laVDzMQW6lrINb',
      dataType: 'json',
      success: function(response) {
        var sheetData = response.content;
        data = parseSheetData(sheetData);
        // 过滤掉全空数据
        displayData = data
  
        displayTable();
      },
      error: function(error) {
        console.error('Error loading file from URL: ' + error);
      }
    });
    $('#uploadButton').show();
    $('#addButton').show();
    $('#searchInput').show();
    $('#refreshButton').show();
    $('#new1').show();
  });
  function parseSheetData(sheetData) {
    var parsedData = [];
    var headers = sheetData[0];
    parsedData.push(headers);
    for (var i = 1; i < sheetData.length; i++) {
      var row = sheetData[i];
      
      var name = row[0];
      var phone = row[1];
      var date = row[2];
      var receipt = row[3];
      var amount = row[4];
      var remark = row[5];
      var donation = row[6];
      if(date!=""){
        var formattedDate = formatDate2(date); // 格式化日期
      }
      var newData = [name, phone, formattedDate, receipt, amount, remark,donation];
      parsedData.push(newData);
    }
    return parsedData;
  }
  //新增按钮
  $('#new1').click(function(){
    $('#new').show();
  });

  // 新增数据
  $('#addButton').click(function() {
      var name = $('#addNameInput').val();
      var phone = $('#addPhoneInput').val();
      var date = $('#addDateInput').val();
      var receipt = $('#addReceiptInput').val();
      var amount = $('#addAmountInput').val();
      var remark = $('#addRemarkInput').val();
      var donation = $('#addJointDonation').val();
  
      if (name.trim() === "") {
        alert("新增数据:姓名不能为空");
        return;
      }
      if (amount !== "" && isNaN(amount)) {
        alert("金额必须是有效的数字");
        return;
      }
      if(date!=""){
        var formattedDate = formatDate1(date); // 格式化日期
      }
      if(amount!=""){
        amount = "RM"+amount; 
      }
      var newData = [name, phone, formattedDate, receipt, amount, remark,donation];
      data.push(newData);
      updateExcelData();
      displayTable();
      $('#addNameInput').val('');
      $('#addPhoneInput').val('');
      $('#addDateInput').val('');
      $('#addReceiptInput').val('');
      $('#addAmountInput').val('');
      $('#addRemarkInput').val('');
      $('#addJointDonation').val('');
      $('#new').hide();
      
  });
  // 格式化日期
  function formatDate1(date) {
    var parts = date.split('-'); // 将日期字符串拆分为年、月、日部分
    var year = parseInt(parts[0]);
    var month = parseInt(parts[1]);
    var day = parseInt(parts[2]);
    var excelDate = Math.floor((Date.UTC(year, month - 1, day) - Date.UTC(1899, 11, 30)) / (24 * 60 * 60 * 1000));
    return excelDate;
  }
  function formatDate2(date) {
    var parts = date.split('-'); // 将日期字符串拆分为年、月、日部分
    var year = parseInt(parts[0]);
    var month = parseInt(parts[1]);
    var day = parseInt(parts[2]);
    var excelDate = Math.floor((Date.UTC(year, month - 1, day+1) - Date.UTC(1899, 11, 30)) / (24 * 60 * 60 * 1000));
    return excelDate;
  }
  



  // 更新Excel数据
  function updateExcelData() {
    var jsonData = JSON.stringify(data); // 将数据转换为JSON字符串
  fetch("https://script.google.com/macros/s/AKfycbyDhl0MYvJfonUAXZpVBIlaNhwFIKmFW0BykpFfkoPVg5mFRBPqYFAZD63TKu0O9DL84Q/exec", {
    method: 'POST',
    cache:"no-cache",
    redirect:"follow",
    body: jsonData,
  })
  .then(function(response) {
    if (response.ok) {
      console.log('Data uploaded successfully!');
    } else {
      throw new Error('Error uploading data: ' + response.status);
    }
  })
  .catch(function(error) {
    console.error(error);
  });
  }

  // 搜索数据
  $('#searchSelect').change(function() {
    var selectedOption = $(this).val();

    // 根据选择的搜索项显示/隐藏相应的输入框
    if (selectedOption === "0" || selectedOption === "1" || selectedOption === "3" || selectedOption === "5") {
      $('#searchInput').show();
      $('#dateInputs').hide();
      $('#amountInput').hide();
    } else if (selectedOption === "2") {
      $('#searchInput').hide();
      $('#dateInputs').show();
      $('#amountInput').hide();
    } else if (selectedOption === "4") {
      $('#searchInput').hide();
      $('#dateInputs').hide();
      $('#amountInput').show();
    }
  });
  var previousSearchData = []; // 用于存储之前的搜索结果

$('#searchButton').click(function() {
  var keyword = $('#searchInput').val().toLowerCase();
  var searchIndex = $('#searchSelect').val();
  var startDate = formatDate1($('#startDateInput').val());
  var endDate = formatDate1($('#endDateInput').val());
  var amountOperator = $('#amountOperator').val();
  var amountValue = parseFloat($('#amountValue').val());
  var usePreviousSearch = $('#usePreviousSearchCheckbox').is(':checked'); // 获取复选框状态
  var matchingIndices = []; // 存储匹配的索引
  var isAmountSearch = searchIndex === "4";
  if (usePreviousSearch && previousSearchData.length > 0&& !isAmountSearch) {
    // 使用之前的搜索结果
    previousSearchData.forEach(function(row, index) {
      if (index === 0) {
        matchingIndices.push(index);
        return;
      }
      if (row[searchIndex] == undefined) {
        var rowValue=""
      }
      else {
        var rowValue = row[searchIndex].toString().toLowerCase();
      }
      var rowDate = rowValue;
      var isMatchKeyword = false;
      var isInRange = true;
      var isAmountMatch = true;
      
      if (searchIndex === "0" || searchIndex === "5") {
        // 对姓名和备注进行关键字匹配
        isMatchKeyword = rowValue.includes(keyword);
      } else if (searchIndex === "1" || searchIndex === "3") {
        // 对手机和收据进行完全匹配
        isMatchKeyword = rowValue === keyword;
      } else if (searchIndex === "2") {
        // 进行日期范围筛选
        isMatchKeyword = true;
        if(!isNaN(startDate) && isNaN(endDate)) {
          isInRange =formatDate(rowDate) === formatDate(startDate)
          var a=formatDate(rowDate)+"\n"+formatDate(startDate)+"\n"+isInRange
        }
        else if (!isNaN(startDate) && !isNaN(endDate)) {
          isInRange = rowDate >= startDate && rowDate <= endDate;
        } else {
          isMatchKeyword = false;
          isInRange = false; // 如果没有有效的日期范围,则不匹配
        }
      } else if (searchIndex === "4") {
        // 进行金额筛选
        if (!isNaN(amountValue)) {
          isMatchKeyword = true;
          var rowAmount = parseFloat(rowValue.replace("rm", ""));
          if (amountOperator === "=") {
            isAmountMatch = rowAmount === amountValue;
          } else if (amountOperator === "<=") {
            isAmountMatch = rowAmount <= amountValue;
          } else if (amountOperator === ">=") {
            isAmountMatch = rowAmount >= amountValue;
          }
        } else {
          isMatchKeyword = false;
          isAmountMatch = false; // 如果没有有效的金额值,则不匹配
        }
      }

      if (isMatchKeyword && isInRange && isAmountMatch) {
        matchingIndices.push(index); // 将匹配的索引添加到 matchingIndices 数组中
      }
    });
    displayData = matchingIndices.map(function(index) {
      return previousSearchData[index]; // 根据索引获取对应的数据
    });
    
  } else {
    // 执行常规的单独搜索

    previousSearchData = [];
    data.forEach(function(row, index) {
      if (index === 0) {
        matchingIndices.push(index);
        return;
      }
      if (row[searchIndex] == undefined) {
        var rowValue=""
      }
      else {
        var rowValue = row[searchIndex].toString().toLowerCase();
      }
      
      var rowDate = rowValue;
      var isMatchKeyword = false;
      var isInRange = true;
      var isAmountMatch = true;
      if (searchIndex === "0" || searchIndex === "5") {
        // 对姓名和备注进行关键字匹配
        isMatchKeyword = rowValue.includes(keyword);
      } else if (searchIndex === "1" || searchIndex === "3") {
        // 对手机和收据进行完全匹配
        isMatchKeyword = rowValue.includes(keyword);
      } else if (searchIndex === "2") {
        // 进行日期范围筛选
        isMatchKeyword = true;
        if(!isNaN(startDate) && isNaN(endDate)) {
          isInRange =formatDate(rowDate) === formatDate(startDate)
          var a=formatDate(rowDate)+"\n"+formatDate(startDate)+"\n"+isInRange
        }
        else if (!isNaN(startDate) && !isNaN(endDate)) {
          isInRange = rowDate >= startDate && rowDate <= endDate;
        } else {
          isMatchKeyword = false;
          isInRange = false; // 如果没有有效的日期范围,则不匹配
        }
      } else if (searchIndex === "4") {
        // 进行金额筛选
        if (!isNaN(amountValue)) {
          isMatchKeyword = true;
          var rowAmount = parseFloat(rowValue.replace("rm", ""));
          if (amountOperator === "=") {
            isAmountMatch = rowAmount === amountValue;
          } else if (amountOperator === "<=") {
            isAmountMatch = rowAmount <= amountValue;
          } else if (amountOperator === ">=") {
            isAmountMatch = rowAmount >= amountValue;
          }
        } else {
          isMatchKeyword = false;
          isAmountMatch = false; // 如果没有有效的金额值,则不匹配
        }
      }

      if (isMatchKeyword && isInRange && isAmountMatch) {
        matchingIndices.push(index); // 将匹配的索引添加到 matchingIndices 数组中
      }
    });
    displayData = matchingIndices.map(function(index) {
      return data[index]; // 根据索引获取对应的数据
    });
  }

  if (!usePreviousSearch) {
    // 更新之前的搜索结果
    previousSearchData = displayData.slice();
  }

  displayTable();

  // 根据搜索结果显示/隐藏上传按钮
  if (
    $('#searchInput').val() === "" &&
    $('#startDateInput').val() === "" &&
    $('#endDateInput').val() === "" &&
    $('#amountValue').val() === ""
  ) {
    $('#uploadButton').show();
  } else {
    $('#uploadButton').hide();
  }
});

  

// 删除数据
$(document).off('click', '.deleteButton').on('click', '.deleteButton', function() {
  var index = $(this).data('index');
  displayData.splice(index, 1);
  data.splice(index, 1); // 从data数组中删除对应数据
  displayTable();
});

$(document).on('click', '.deleteButton', function() {
  var index = $(this).data('index');
  displayData.splice(index, 1);
  data.splice(index, 1); // 从data数组中删除对应数据
  displayTable();
});

var nameSortOrder = 1; // 跟踪排序顺序,初始值为1

// 按姓名排序
$('#sortNameButton').click(function() {
  displayData.sort(function(a, b) {
    if (a === displayData[0] || b === displayData[0]) {
      return 0; // 跳过第一行的排序
    }
    return a[0].localeCompare(b[0]) * nameSortOrder;
  });

  nameSortOrder *= -1; // 切换排序顺序

  data = displayData;
  displayTable();
});


var dateSortOrder = 1; // 跟踪排序顺序,初始值为1

// 按日期排序
$('#sortDateButton').click(function() {
  displayData.sort(function(a, b) {
    if (a === displayData[0] || b === displayData[0]) {
      return 0; // 跳过第一行的排序
    }
    if (a[2] === undefined) {
      return 1; // 将 undefined 放在最后
    }
    if (b[2] === undefined) {
      return -1; // 将 undefined 放在最后
    }
    if (new Date(a[2]) < new Date(b[2])) {
      return -1 * dateSortOrder; // 小到大排序
    }
    if (new Date(a[2]) > new Date(b[2])) {
      return 1 * dateSortOrder; // 大到小排序
    }
    return 0;
  });

  dateSortOrder *= -1; // 切换排序顺序

  data = displayData;
  displayTable();
});


var amountSortOrder = 1; // 跟踪排序顺序,初始值为1

// 按金额排序
$('#sortAmountButton').click(function() {
  displayData.sort(function(a, b) {
    if (a === displayData[0] || b === displayData[0]) {
      return 0; // 跳过第一行的排序
    }
    var amountA = parseFloat(a[4].replace('RM', ''));
    var amountB = parseFloat(b[4].replace('RM', ''));

    if (amountA < amountB) {
      return -1 * amountSortOrder; // 小到大排序
    }
    if (amountA > amountB) {
      return 1 * amountSortOrder; // 大到小排序
    }
    return 0;
  });

  amountSortOrder *= -1; // 切换排序顺序

  data = displayData;
  displayTable();
});



// 刷新表格
$('#refreshButton').click(function() {
  previousSearchData = [];
  displayData = data.slice();
  displayTable();
  $('#addNameInput').val('');
  $('#addPhoneInput').val('');
  $('#addDateInput').val('');
  $('#addReceiptInput').val('');
  $('#addAmountInput').val('');
  $('#addRemarkInput').val('');
  $('#addJointDonation').val('');
  $('#uploadButton').show();
  $('#new').hide();
  $('#usePreviousSearchCheckbox').prop('checked', false);
  
});


// 上传数据到Excel
$('#uploadButton').click(function() {
  var jsonData = JSON.stringify(data); // 将数据转换为JSON字符串
  fetch("https://script.google.com/macros/s/AKfycbyDhl0MYvJfonUAXZpVBIlaNhwFIKmFW0BykpFfkoPVg5mFRBPqYFAZD63TKu0O9DL84Q/exec", {
    method: 'POST',
    cache:"no-cache",
    redirect:"follow",
    body: jsonData,
  })
  .then(function(response) {
    if (response.ok) {
      console.log('Data uploaded successfully!');
    } else {
      throw new Error('Error uploading data: ' + response.status);
    }
  })
  .catch(function(error) {
    console.error(error);
  });
  alert("更新成功")
});



// 显示数据到表格
function displayTable() {
  var tableBody = $('#dataTable tbody');
  tableBody.empty();
  for (var i = 1; i < displayData.length; i++) {
    var row = '<tr>';
    for (var j = 0; j < 7; j++) {
      var cellData = "";
      if (displayData[i].length > j) {
        cellData = displayData[i][j];
        if (cellData == undefined) cellData = "";
        if (j == 2 && cellData != "") {
          cellData = formatDate(cellData);
        }
        if (j == 6 && cellData != "") {
          // = "合捐";
          cellData='<button class="combineButton" data-index="' + i + '" style="background-color:#a8badb;color:black;border:none;font-weight:bold;;font-size:15px;">合捐</button>';
        }
      }
      row += '<td>' + cellData + '</td>';
    }
    row += '<td><button class="editButton " data-index="' + i + '">修改</button> <button class="deleteButton" data-index="' + i + '">删除</button></td>';
    row += '</tr>';
    tableBody.append(row);
  }

 // 合捐按钮点击事件处理程序
$('.combineButton').click(function() {
  var dataIndex = $(this).data('index');
  var combinedData = displayData[dataIndex];
  var names = combinedData[6].split(","); // 分割合捐名字
  var message = "合捐名单:\n";
  message += names.join("\n"); // 使用换行符连接名字
  alert(message);
});

}


// 格式化日期
function formatDate(date) {
  var excelDate = parseInt(date); // 将日期字符串转换为整数
  var jsDate = new Date((excelDate - 1) * 24 * 60 * 60 * 1000 + Date.UTC(1900, 0, 0));
  var year = jsDate.getFullYear();
  var month = jsDate.getMonth() + 1;
  var day = jsDate.getDate();
  return year + '/' + month + '/' + day;
}


// 修改单个资料
$(document).on('click', '.editButton', function() {
var index = $(this).data('index');
var oldValue = displayData[index];
var num = prompt('要修改什么: [0]姓名 , [1]电话号码 , [2]日期 , [3]收据 , [4]金额 , [5]备注 , [6]合捐');
switch(num) {
  case '0':
      var newname = prompt('姓名:',oldValue[0]);
      oldValue[num]=newname;
      break;
  case '1':
      var newphone = prompt('手机:',oldValue[1]);
      oldValue[num]=newphone;
      break;
  case '2':
      var newdate = prompt('日期:',oldValue[2]);
      oldValue[num]=newdate;
      break;
  case '3':
      var newreceipt = prompt('收据:',oldValue[3]);
      oldValue[num]=newreceipt;
      break;
  case '4':
      var newamount = prompt('金额:',oldValue[4]);
      oldValue[num]="RM"+newamount;
      break;
  case '5':
      var newremark = prompt('备注:',oldValue[5]);
      oldValue[num]=newremark;
      break;
  case '6':
      var newremark = prompt('合捐:',oldValue[6]);
      oldValue[num]=newremark;
      break;
  default:
    // code block
}
if (oldValue !== null) {
displayData[index] = oldValue;
data[index] = oldValue; // 更新data数组中的对应数据
displayTable();
}
});

// 将二进制字符串转换为字节数组
function s2ab(s) {
  var buf = new ArrayBuffer(s.length);
  var view = new Uint8Array(buf);
  for (var i = 0; i < s.length; i++) {
    view[i] = s.charCodeAt(i) & 0xFF;
  }
  return buf;
}


});

