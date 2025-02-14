function refreshPageAccessToken() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Settings");
    sheet.getRange("A1").setValue("Access Token");
  }

  var oldAccessToken = sheet.getRange("A2").getValue(); // Token hiện tại
  var appId = "your_app_id"; // App ID
  var appSecret = "your_app_secret"; // App Secret
  sheet.getRange("A4").setValue(oldAccessToken);
  // Hàm hủy quyền token cũ
function revokeOldAccessToken(oldAccessToken) {
  var url = "https://graph.facebook.com/v22.0/<user id trong ứng dụng của bạn>/permissions?access_token=" + oldAccessToken;

  try {
    // Yêu cầu hủy quyền (revoke) token
    var response = UrlFetchApp.fetch(url, {
      method: "DELETE"  // Sử dụng phương thức DELETE để yêu cầu hủy quyền
    });

    var jsonResponse = JSON.parse(response.getContentText());
    Logger.log("Token cũ đã được hủy quyền: " + JSON.stringify(jsonResponse));

  } catch (e) {
    Logger.log("Lỗi khi hủy token cũ: " + e.toString());
  }
}
  // Hủy quyền token cũ
  revokeOldAccessToken(oldAccessToken);

  // URL để lấy token mới
  var url = "https://graph.facebook.com/v22.0/oauth/access_token" +
            "?grant_type=fb_exchange_token" +
            "&client_id=" + appId +
            "&client_secret=" + appSecret +
            "&fb_exchange_token=" + oldAccessToken;

  try {
    var response = UrlFetchApp.fetch(url);
    var json = JSON.parse(response.getContentText());

    if (json.access_token) {
      sheet.getRange("A2").setValue(json.access_token); // Lưu token mới vào A2
      Logger.log("Token mới đã được cập nhật!");
    } else {
      Logger.log("Lỗi lấy token mới: " + response.getContentText());
    }
  } catch (e) {
    Logger.log("Lỗi API: " + e.toString());
  }
}
// Hàm kiểm tra tính hợp lệ của token sau khi hủy
function checkTokenValidity(oldAccessToken) {
  var url = "https://graph.facebook.com/debug_token?input_token=" + oldAccessToken + "&access_token=" + oldAccessToken;

  try {
    var response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true // Bỏ qua lỗi HTTP và lấy phản hồi chi tiết
    });
    var jsonResponse = JSON.parse(response.getContentText());

    if (jsonResponse.data && jsonResponse.data.is_valid) {
      Logger.log("Token vẫn còn hợp lệ.");
    } else {
      Logger.log("Token đã hết hạn hoặc không hợp lệ. "+oldAccessToken);
    }
  } catch (e) {
    Logger.log("Lỗi khi kiểm tra token: " + e.toString());
  }
}
function tokenHientai() {
  Logger.log("Token hiện tại: " + getStoredToken());
}
function getStoredToken() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  return sheet.getRange("A2").getValue();
}
function postConfessionsToFacebook() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Câu trả lời biểu mẫu"); // Thay "Sheet1" bằng tên sheet của bạn
  var e2Formula = sheet.getRange("E2").getFormula(); // Lấy công thức trong ô E2
  var match = e2Formula.match(/C(\d+)/i); // Tìm số hàng mà E2 đang tham chiếu đến

  if (!match) {
    Logger.log("Không tìm thấy ô tham chiếu trong E2.");
    return;
  }

  var startRow = parseInt(match[1]); // Lấy số hàng từ công thức trong E2
  var endRow = startRow + 9; // Lấy 20 dòng từ Cx -> C(x+19)
  var values = sheet.getRange("C" + startRow + ":C" + endRow).getValues(); // Lấy dữ liệu

  // Kiểm tra xem có dữ liệu không
  if (values.flat().join("").trim() === "") {
    Logger.log("Không có confession nào để đăng.");
    return;
  }

  // Định dạng nội dung bài đăng
  var postContent = "🌸💮🌸💮🌸\n";
  values.forEach(function(row, index) {
    if (row[0].trim() !== "") { // Bỏ qua dòng trống
      postContent += "#cfs" + (startRow + index + 997) + " " + row[0] + "\n";
    }
  });
  postContent += "🌸💮🌸💮🌸";

  // Thông tin Facebook API
  var pageAccessToken = getStoredToken(); // Token
  var pageId = "Page ID của bạn"; // ID Page
  var url = "https://graph.facebook.com/" + pageId + "/feed";

  // Gửi yêu cầu đăng bài lên Facebook
  var options = {
    method: "post",
    payload: {
      message: postContent,
      access_token: pageAccessToken
    }
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());

    if (json.id) {
      Logger.log("Đăng bài thành công! ID bài viết: " + json.id);

      // Cập nhật công thức trong E2 để tham chiếu đến ô C(x+10)
      var nextRow = startRow + 10;
      sheet.getRange("E2").setFormula("=C" + nextRow);
    } else {
      Logger.log("Lỗi khi đăng bài: " + response.getContentText());
    }
  } catch (e) {
    // Kiểm tra nếu lỗi là do token hết hạn (code 190)
    if (e.message.includes('Error validating access token: Session has expired on')) {
      Logger.log("Token hết hạn, đang làm mới token...");

      // Gọi hàm refreshPageAccessToken để lấy token mới
      refreshPageAccessToken();

      // Thử lại sau khi token đã được làm mới
      postConfessionsToFacebook();
    } else {
      Logger.log("Lỗi không phải do token: " + e.toString());
    }
  }
}


/* function checkMonthAndRun() {
  var currentMonth = new Date().getMonth(); // Lấy tháng hiện tại (0-11, 0 là tháng 1)
  
  // Kiểm tra nếu là tháng chẵn (tháng 2, 4, 6, 8, 10, 12)
  if ([1, 3, 5, 7, 9, 11].indexOf(currentMonth) !== -1) {
    // Gọi hàm refreshPageAccessToken nếu là tháng chẵn
    refreshPageAccessToken();
  } else {
    Logger.log("Không phải tháng chẵn. Trigger không thực hiện.");
  }
}
function createTimeDrivenTrigger() {
  // Thêm trigger chạy mỗi tháng vào ngày mùng 4 lúc 9 giờ sáng
  ScriptApp.newTrigger('checkMonthAndRun') 
      .timeBased()
      .onMonthDay(4)  // Đặt trigger chạy vào ngày mùng 4 mỗi tháng
      .atHour(3)      // Thời gian chạy là 3 giờ sáng
      .create();
} */

function postConfessionsToInstagram() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Câu trả lời biểu mẫu");
  var e2Formula = sheet.getRange("F2").getFormula();
  var match = e2Formula.match(/C(\d+)/i);
  
  if (!match) {
    Logger.log("Không tìm thấy ô tham chiếu trong F2.");
    return;
  }
  
  var startRow = parseInt(match[1]);
  var endRow = startRow + 9;
  var values = sheet.getRange("C" + startRow + ":C" + endRow).getValues();
  
  if (values.flat().join("").trim() === "") {
    Logger.log("Không có confession nào để đăng.");
    return;
  }
  
  var today = new Date();
  var day = today.getDate();
  var month = today.getMonth() + 1; // Tháng trong JS bắt đầu từ 0 nên cần +1
  var year = today.getFullYear();
  var formattedDate = "Confessions ngày " + day + " tháng " + month + " năm " + year + "\n";

  var postContent = formattedDate;
  values.forEach(function(row, index) {
    if (row[0].trim() !== "") {
      postContent += "\n#cfs" + (startRow + index + 997) + "\n" + row[0] + "\n";
    }
  });
  postContent += "🌸💮🌸💮🌸";
  
  var imgurImageUrl = "https://i.imgur.com/caigiday"; // Thay bằng link ảnh Imgur thật
  var pageAccessToken = getStoredToken();
  var instagramAccountId = "YOUR_ID"; //Thay bằng Instagram Business ID của bạn
  
  var createMediaUrl = "https://graph.facebook.com/v22.0/" + instagramAccountId + "/media";
  var createMediaOptions = {
    method: "post",
    payload: {
      image_url: imgurImageUrl,
      caption: postContent,
      access_token: pageAccessToken
    }
  };
  
  try {
    var response = UrlFetchApp.fetch(createMediaUrl, createMediaOptions);
    var json = JSON.parse(response.getContentText());
    
    if (json.id) {
      var publishUrl = "https://graph.facebook.com/v22.0/" + instagramAccountId + "/media_publish";
      var publishOptions = {
        method: "post",
        payload: {
          creation_id: json.id,
          access_token: pageAccessToken
        }
      };
      
      var publishResponse = UrlFetchApp.fetch(publishUrl, publishOptions);
      var publishJson = JSON.parse(publishResponse.getContentText());
      
      if (publishJson.id) {
        Logger.log("Đăng bài thành công! ID bài viết: " + publishJson.id);
        var nextRow = startRow + 10;
        sheet.getRange("F2").setFormula("=C" + nextRow);
      } else {
        Logger.log("Lỗi khi xuất bản bài: " + publishResponse.getContentText());
      }
    } else {
      Logger.log("Lỗi khi tạo media: " + response.getContentText());
    }
  } catch (e) {
    Logger.log("Lỗi: " + e.toString());
  }
}
