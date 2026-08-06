/**
 * Google Apps Script for Brigade Granada Landing Page Form Submission
 */

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // If the sheet is empty, initialize headers
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Timestamp", "Name", "Email", "Phone", "Form Source", "Campaign", "Source", "Medium", "Keyword", "GCLID", "Landing Page", "Browser"]);
      
      // Style headers: bold, light gray background, freeze header row
      sheet.getRange("A1:L1").setFontWeight("bold").setBackground("#f3f3f3");
      sheet.setFrozenRows(1);
    }
    
    // Parse the incoming JSON data
    var data = JSON.parse(e.postData.contents);
    
    var timestamp = new Date();
    var name = data.name || "N/A";
    var email = data.email || "N/A";
    var phone = data.phone || "N/A";
    var source = data.source || "Unknown";
    
    // Captured tracking fields
    var campaign = data.campaign || "";
    var sourceUtm = data.utm_source || data.source_utm || "";
    var medium = data.medium || "";
    var keyword = data.keyword || "";
    var gclid = data.gclid || "";
    var landingPage = data.landing_page || "";
    var browser = data.browser || "";
    
    // Prepend a single quote if the phone number starts with '+' to prevent Google Sheets from parsing it as a formula
    if (phone.toString().trim().indexOf('+') === 0) {
      phone = "'" + phone.toString().trim();
    }
    
    // Append the row to Google Sheets
    sheet.appendRow([timestamp, name, email, phone, source, campaign, sourceUtm, medium, keyword, gclid, landingPage, browser]);
    
    // Send clean HTML email alert to Realhubbsales@gmail.com
    try {
      var emailRecipient = "Realhubbsales@gmail.com";
      var emailSubject = "🏡 Brigade Granada: New Lead - " + name;
      
      // Clean up single quote from phone display in the email
      var displayPhone = phone.toString().trim().replace(/^'/, '');
      var cleanPhoneForWhatsApp = displayPhone.replace(/\s+/g, '').replace('+', '');
      if (!cleanPhoneForWhatsApp.startsWith('91') && cleanPhoneForWhatsApp.length === 10) {
        cleanPhoneForWhatsApp = '91' + cleanPhoneForWhatsApp;
      }
      
      // Beautiful HTML Email Template
      var emailHtmlBody = 
        "<div style='font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.05);'>" +
          "<div style='background-color: #00274D; color: #ffffff; padding: 24px; text-align: center;'>" +
            "<h2 style='margin: 0; font-size: 22px; font-weight: 600; letter-spacing: 1px;'>BRIGADE GRANADA</h2>" +
            "<p style='margin: 5px 0 0 0; font-size: 13px; color: #D4AF37; text-transform: uppercase; font-weight: bold; letter-spacing: 0.5px;'>New Website Lead</p>" +
          "</div>" +
          "<div style='padding: 24px; background-color: #ffffff; color: #333333;'>" +
            "<p style='font-size: 15px; line-height: 1.5; margin-top: 0;'>Hello Team,</p>" +
            "<p style='font-size: 15px; line-height: 1.5;'>A new customer interest has been captured from the <strong>Brigade Granada</strong> landing page:</p>" +
            
            "<table style='width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 14px;'>" +
              "<tr style='background-color: #f9f9f9;'>" +
                "<td style='padding: 12px; font-weight: bold; width: 130px; border-bottom: 1px solid #eeeeee;'>Customer Name</td>" +
                "<td style='padding: 12px; border-bottom: 1px solid #eeeeee; color: #00274D; font-weight: bold;'>" + name + "</td>" +
              "</tr>" +
              "<tr>" +
                "<td style='padding: 12px; font-weight: bold; border-bottom: 1px solid #eeeeee;'>Phone Number</td>" +
                "<td style='padding: 12px; border-bottom: 1px solid #eeeeee; font-weight: bold;'>" + displayPhone + "</td>" +
              "</tr>" +
              "<tr style='background-color: #f9f9f9;'>" +
                "<td style='padding: 12px; font-weight: bold; border-bottom: 1px solid #eeeeee;'>Email Address</td>" +
                "<td style='padding: 12px; border-bottom: 1px solid #eeeeee; color: #666666;'>" + email + "</td>" +
              "</tr>" +
              "<tr>" +
                "<td style='padding: 12px; font-weight: bold; border-bottom: 1px solid #eeeeee;'>Form Source</td>" +
                "<td style='padding: 12px; border-bottom: 1px solid #eeeeee;'><span style='background-color: #e8f4fd; color: #00274D; padding: 4px 8px; border-radius: 6px; font-size: 12px; font-weight: bold;'>" + source + "</span></td>" +
              "</tr>" +
              "<tr style='background-color: #f9f9f9;'>" +
                "<td style='padding: 12px; font-weight: bold; border-bottom: 1px solid #eeeeee;'>Campaign</td>" +
                "<td style='padding: 12px; border-bottom: 1px solid #eeeeee; color: #666666;'>" + (campaign || "N/A") + "</td>" +
              "</tr>" +
              "<tr>" +
                "<td style='padding: 12px; font-weight: bold; border-bottom: 1px solid #eeeeee;'>Source (UTM)</td>" +
                "<td style='padding: 12px; border-bottom: 1px solid #eeeeee; color: #666666;'>" + (sourceUtm || "N/A") + "</td>" +
              "</tr>" +
              "<tr style='background-color: #f9f9f9;'>" +
                "<td style='padding: 12px; font-weight: bold; border-bottom: 1px solid #eeeeee;'>Medium</td>" +
                "<td style='padding: 12px; border-bottom: 1px solid #eeeeee; color: #666666;'>" + (medium || "N/A") + "</td>" +
              "</tr>" +
              "<tr>" +
                "<td style='padding: 12px; font-weight: bold; border-bottom: 1px solid #eeeeee;'>Keyword</td>" +
                "<td style='padding: 12px; border-bottom: 1px solid #eeeeee; color: #666666;'>" + (keyword || "N/A") + "</td>" +
              "</tr>" +
              "<tr style='background-color: #f9f9f9;'>" +
                "<td style='padding: 12px; font-weight: bold; border-bottom: 1px solid #eeeeee;'>GCLID</td>" +
                "<td style='padding: 12px; border-bottom: 1px solid #eeeeee; color: #666666;'>" + (gclid || "N/A") + "</td>" +
              "</tr>" +
              "<tr>" +
                "<td style='padding: 12px; font-weight: bold; border-bottom: 1px solid #eeeeee;'>Landing Page</td>" +
                "<td style='padding: 12px; border-bottom: 1px solid #eeeeee; color: #666666;'>" + (landingPage || "N/A") + "</td>" +
              "</tr>" +
              "<tr style='background-color: #f9f9f9;'>" +
                "<td style='padding: 12px; font-weight: bold; border-bottom: 1px solid #eeeeee;'>Browser</td>" +
                "<td style='padding: 12px; border-bottom: 1px solid #eeeeee; color: #666666;'>" + (browser || "N/A") + "</td>" +
              "</tr>" +
              "<tr>" +
                "<td style='padding: 12px; font-weight: bold; border-bottom: 1px solid #eeeeee;'>Time Captured</td>" +
                "<td style='padding: 12px; border-bottom: 1px solid #eeeeee; color: #888888;'>" + timestamp.toLocaleString("en-IN") + "</td>" +
              "</tr>" +
            "</table>" +
            
            "<div style='margin-top: 28px; text-align: center;'>" +
              "<a href='tel:" + displayPhone.replace(/\s+/g, '') + "' style='display: inline-block; background-color: #00274D; color: #ffffff; text-decoration: none; padding: 12px 22px; margin: 5px; border-radius: 8px; font-weight: bold; font-size: 14px;'>📞 Call Client</a>" +
              "<a href='https://wa.me/" + cleanPhoneForWhatsApp + "' target='_blank' style='display: inline-block; background-color: #25D366; color: #ffffff; text-decoration: none; padding: 12px 22px; margin: 5px; border-radius: 8px; font-weight: bold; font-size: 14px;'>💬 WhatsApp Client</a>" +
            "</div>" +
          "</div>" +
          "<div style='background-color: #f5f5f5; color: #777777; padding: 16px; text-align: center; font-size: 11px; border-top: 1px solid #e0e0e0;'>" +
            "This lead was captured from the Brigade Granada Microsite and saved to Google Sheets.<br/>" +
            "&copy; Realhubb Ventures Pvt Ltd." +
          "</div>" +
        "</div>";

      MailApp.sendEmail({
        to: emailRecipient,
        subject: emailSubject,
        htmlBody: emailHtmlBody
      });
    } catch (emailErr) {
      console.error("Email sending failed: " + emailErr.toString());
    }
    
    // Return a success JSON response
    return ContentService
      .createTextOutput(JSON.stringify({ "result": "success" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    // Return error message if something goes wrong
    return ContentService
      .createTextOutput(JSON.stringify({ "result": "error", "error": err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Add support for GET request test
function doGet(e) {
  return ContentService.createTextOutput("Web App is running! Use POST to submit form data.");
}

function testEmail() {
  // Simulating details for a test email
  var timestamp = new Date();
  var name = "Test Lead (Granada)";
  var email = "test@example.com";
  var phone = "+91 9999999999";
  var source = "Test Verification";
  var campaign = "test_campaign";
  var sourceUtm = "google";
  var medium = "cpc";
  var keyword = "luxury apartments";
  var gclid = "test_gclid_123";
  var landingPage = "https://example.com/granada";
  var browser = "Chrome";
  
  var emailRecipient = "Realhubbsales@gmail.com";
  var emailSubject = "🏡 Brigade Granada: New Lead - " + name;
  var displayPhone = phone.replace(/^'/, '');
  var cleanPhoneForWhatsApp = displayPhone.replace(/\s+/g, '').replace('+', '');
  
  var emailHtmlBody = 
    "<div style='font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.05);'>" +
      "<div style='background-color: #00274D; color: #ffffff; padding: 24px; text-align: center;'>" +
        "<h2 style='margin: 0; font-size: 22px; font-weight: 600; letter-spacing: 1px;'>BRIGADE GRANADA</h2>" +
        "<p style='margin: 5px 0 0 0; font-size: 13px; color: #D4AF37; text-transform: uppercase; font-weight: bold; letter-spacing: 0.5px;'>New Website Lead</p>" +
      "</div>" +
      "<div style='padding: 24px; background-color: #ffffff; color: #333333;'>" +
        "<p style='font-size: 15px; line-height: 1.5; margin-top: 0;'>Hello Team,</p>" +
        "<p style='font-size: 15px; line-height: 1.5;'>A new customer interest has been captured from the <strong>Brigade Granada</strong> landing page:</p>" +
        
        "<table style='width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 14px;'>" +
          "<tr style='background-color: #f9f9f9;'>" +
            "<td style='padding: 12px; font-weight: bold; width: 130px; border-bottom: 1px solid #eeeeee;'>Customer Name</td>" +
            "<td style='padding: 12px; border-bottom: 1px solid #eeeeee; color: #00274D; font-weight: bold;'>" + name + "</td>" +
          "</tr>" +
          "<tr>" +
            "<td style='padding: 12px; font-weight: bold; border-bottom: 1px solid #eeeeee;'>Phone Number</td>" +
            "<td style='padding: 12px; border-bottom: 1px solid #eeeeee; font-weight: bold;'>" + displayPhone + "</td>" +
          "</tr>" +
          "<tr style='background-color: #f9f9f9;'>" +
            "<td style='padding: 12px; font-weight: bold; border-bottom: 1px solid #eeeeee;'>Email Address</td>" +
            "<td style='padding: 12px; border-bottom: 1px solid #eeeeee; color: #666666;'>" + email + "</td>" +
          "</tr>" +
          "<tr>" +
            "<td style='padding: 12px; font-weight: bold; border-bottom: 1px solid #eeeeee;'>Form Source</td>" +
            "<td style='padding: 12px; border-bottom: 1px solid #eeeeee;'><span style='background-color: #e8f4fd; color: #00274D; padding: 4px 8px; border-radius: 6px; font-size: 12px; font-weight: bold;'>" + source + "</span></td>" +
          "</tr>" +
          "<tr style='background-color: #f9f9f9;'>" +
            "<td style='padding: 12px; font-weight: bold; border-bottom: 1px solid #eeeeee;'>Campaign</td>" +
            "<td style='padding: 12px; border-bottom: 1px solid #eeeeee; color: #666666;'>" + campaign + "</td>" +
          "</tr>" +
          "<tr>" +
            "<td style='padding: 12px; font-weight: bold; border-bottom: 1px solid #eeeeee;'>Source (UTM)</td>" +
            "<td style='padding: 12px; border-bottom: 1px solid #eeeeee; color: #666666;'>" + sourceUtm + "</td>" +
          "</tr>" +
          "<tr style='background-color: #f9f9f9;'>" +
            "<td style='padding: 12px; font-weight: bold; border-bottom: 1px solid #eeeeee;'>Medium</td>" +
            "<td style='padding: 12px; border-bottom: 1px solid #eeeeee; color: #666666;'>" + medium + "</td>" +
          "</tr>" +
          "<tr>" +
            "<td style='padding: 12px; font-weight: bold; border-bottom: 1px solid #eeeeee;'>Keyword</td>" +
            "<td style='padding: 12px; border-bottom: 1px solid #eeeeee; color: #666666;'>" + keyword + "</td>" +
          "</tr>" +
          "<tr style='background-color: #f9f9f9;'>" +
            "<td style='padding: 12px; font-weight: bold; border-bottom: 1px solid #eeeeee;'>GCLID</td>" +
            "<td style='padding: 12px; border-bottom: 1px solid #eeeeee; color: #666666;'>" + gclid + "</td>" +
          "</tr>" +
          "<tr>" +
            "<td style='padding: 12px; font-weight: bold; border-bottom: 1px solid #eeeeee;'>Landing Page</td>" +
            "<td style='padding: 12px; border-bottom: 1px solid #eeeeee; color: #666666;'>" + landingPage + "</td>" +
          "</tr>" +
          "<tr style='background-color: #f9f9f9;'>" +
            "<td style='padding: 12px; font-weight: bold; border-bottom: 1px solid #eeeeee;'>Browser</td>" +
            "<td style='padding: 12px; border-bottom: 1px solid #eeeeee; color: #666666;'>" + browser + "</td>" +
          "</tr>" +
          "<tr>" +
            "<td style='padding: 12px; font-weight: bold; border-bottom: 1px solid #eeeeee;'>Time Captured</td>" +
            "<td style='padding: 12px; border-bottom: 1px solid #eeeeee; color: #888888;'>" + timestamp.toLocaleString("en-IN") + "</td>" +
          "</tr>" +
        "</table>" +
        
        "<div style='margin-top: 28px; text-align: center;'>" +
          "<a href='tel:" + displayPhone.replace(/\s+/g, '') + "' style='display: inline-block; background-color: #00274D; color: #ffffff; text-decoration: none; padding: 12px 22px; margin: 5px; border-radius: 8px; font-weight: bold; font-size: 14px;'>📞 Call Client</a>" +
          "<a href='https://wa.me/" + cleanPhoneForWhatsApp + "' target='_blank' style='display: inline-block; background-color: #25D366; color: #ffffff; text-decoration: none; padding: 12px 22px; margin: 5px; border-radius: 8px; font-weight: bold; font-size: 14px;'>💬 WhatsApp Client</a>" +
        "</div>" +
      "</div>" +
      "<div style='background-color: #f5f5f5; color: #777777; padding: 16px; text-align: center; font-size: 11px; border-top: 1px solid #e0e0e0;'>" +
        "This lead was captured from the Brigade Granada Microsite and saved to Google Sheets.<br/>" +
        "&copy; Realhubb Ventures Pvt Ltd." +
      "</div>" +
    "</div>";

  MailApp.sendEmail({
    to: emailRecipient,
    subject: emailSubject,
    htmlBody: emailHtmlBody
  });
}
