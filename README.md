  $.ajax({url: "/getReceipts.json", data: {role : "report", start : firstDate, end : lastDate, centerID: centerID}, async: false, dataType: "json", success: function(data){
        if (data.success) {
          if(data.datas) process.getReceiptPlMapData(data.datas);
          plMapData["params"] = {"filename": centerName + "(" + firstDate + "-" + lastDate + ")", "centerName": centerName}
          postData = JSON.stringify(plMapData);
 
          $.ajax({url: "http://www.azhibox.com/index.php?type=pl", type: "post", data: {data: postData}, async: false, dataType: "text", success: function(fileURL){
            $("#exportPL").text("导出PL表");
            $("#exportPL").prop("disabled", false);
 
            $("#exportPL").after('<a target="_blank" href="' + fileURL + '" id="plFileUrl"></a>');
            document.getElementById("plFileUrl").click(); 
          }});
        } else {
          alert('获取报销数据出错！');
        }   
      }});
