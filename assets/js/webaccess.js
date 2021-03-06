var width = $(window).width(),
    host = 'https://yutinghcw.github.io/edm',
    initial = host + "/initial/webaccess.html",
    currentYear = new Date().getFullYear();

//修改目的：提供EDM公版//
//修改1：將電子報版頭資訊改到表單中填入，也讓Speadsheet文件表格單純化//
//修改2：簡化表單，只保留填入Google文件序號欄位，其它都在spreadsheet中編輯//
//修改3：資料表格增加「按鈕文字」欄位//-->
//修改4：調整工具的樣式與SOP，新增多欄版型判斷頭、尾的邏輯//

//從參數製作html原始碼
function makeSourceCode(key, worksheet) {
    var codeHead = "",
        previewHead = "",
        codeBody = "",
        trackFoot = "",
        previewFoot = "",
        codeFoot = "";
    var gspreadsheets = "https://spreadsheets.google.com/feeds/list/" + key + "/" + worksheet + "/public/values?alt=json&sq=%E9%A1%AF%E7%A4%BA=y";

    //以Json格式載入文件
    var data = {
        funcColum: []
    };
    $.getJSON(gspreadsheets, function(l) {
        //將載入列表轉換為清單的格式
        $.each(l.feed.entry, function(i, f) {
            data.funcColum[i] = {
                columType: f.gsx$欄型.$t,
                columLink: f.gsx$連結.$t,
                columImage: f.gsx$圖片.$t,
                columChannel: f.gsx$小帽.$t,
                columTitle: f.gsx$標題.$t,
                columText: f.gsx$內文.$t
            };
        });

        //組裝原始碼第一段，版頭及表頭
        codeHead = '<!doctype html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><meta http-equiv="X-UA-Compatible" content="IE=edge"><title>EmailTemplate-Responsive</title><meta content="telephone=no" name="format-detection"><link rel="preconnect" href="https://fonts.gstatic.com"><link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500&family=Roboto:wght@300;400;500&display=swap" rel="stylesheet"><link rel="shortcut icon" href="https://www.cw.com.tw/assets_new/img/favicon.ico" type="image/x-icon" /><style type="text/css">html,body{margin:0 !important;padding:0 !important;height:100% !important;width:100% !important}*{-ms-text-size-adjust:100%;-webkit-text-size-adjust:100%}.ExternalClass{width:100%}div[style*="margin: 16px 0"]{margin:0 !important}table,td{mso-table-lspace:0pt !important;mso-table-rspace:0pt !important}table{border-spacing:0 !important;border-collapse:collapse !important;table-layout:fixed !important;margin:0 auto !important}table table table{table-layout:auto}img{-ms-interpolation-mode:bicubic}.yshortcuts a{border-bottom:none !important}a[x-apple-data-detectors]{color:inherit !important}</style><style type="text/css">.button-td,.button-a{transition:all 100ms ease-in}.button-td:hover,.button-a:hover{background:#555 !important;border-color:#555 !important}@media screen and (max-width: 700px){.email-container{width:100% !important}.fluid,.fluid-centered{max-width:100% !important;height:auto !important;margin-left:auto !important;margin-right:auto !important}.fluid-centered{margin-left:auto !important;margin-right:auto !important}.stack-column,.stack-column-center{display:block !important;width:100% !important;max-width:100% !important;direction:ltr !important}.stack-column-center{text-align:center !important}.center-on-narrow{text-align:center !important;display:block !important;margin-left:auto !important;margin-right:auto !important;float:none !important}table.center-on-narrow{display:inline-block !important}}</style></head><body bgcolor="#fff" width="100%" style="margin: 0;" yahoo="yahoo">';

        previewHead = '<table bgcolor="#fff" cellpadding="0" cellspacing="0" border="0" height="100%" width="100%" style="border-collapse:collapse;"><tr><td><center style="width: 100%;"><div style="display: none; font-size: 1px; line-height: 1px; max-height: 0px; max-width: 0px; opacity: 0; overflow: hidden; mso-hide: all; font-family: sans-serif;">' + data.funcColum[0].columText + '</div><table align="center" width="700" class="email-container"><tr bgcolor="#d60c18"><td style="padding: 1rem 2.857143% 0;"><a style="display: block; line-height: 0;" href="https://www.cw.com.tw/" target="_blank" rel="noopener noreferrer"><img src="https://www.cw.com.tw/assets_new/img/logo.png" style="display: inline-block; height: 30px;" alt="天下雜誌"></a></td></tr><tr align="right" bgcolor="#d60c18"><td style="padding: 30px 2.857143% 0; color: #fff; font-size: 2rem; letter-spacing: 1px; line-height: 1.5;">訂戶週報</td></tr><tr align="right" bgcolor="#d60c18"><td style="padding: 0 2.857143% 1rem; color: #fff; font-size: 1rem; letter-spacing: 3px; line-height: 1.5;">' + data.funcColum[0].columLink + '</td></tr><tr class="unsubscribe"><td style="padding: 10px 2.857143%; text-align: right; border-bottom: 1px solid #c9c9c9;"><p style="margin: 0; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif;"><a href="' + initial + "?file=" + key + '&email=%%email%%" target="_blank" rel="noopener noreferrer" style="color:#999; font-size: 0.8125rem; text-decoration: none;">內容若無法正常顯示，請點我</a></p></td></tr></table><table cellspacing="0" cellpadding="0" border="0" align="center" bgcolor="#fff" width="700" class="email-container">';

        //組裝原始碼第二段，填入spreadsheet的內容部份
        var multiColumCont = 0;
        $(function() {
            $.each(data.funcColum, function(i, f) {
                if (i == 0) {
                    return true;
                }
                switch (f.columType) {
                    case 'head':
                        console.log('head');
                        break;
                    case 'Single-Image':
                        console.log('Single-Image, ' + multiColumCont);
                        codeBody += checkColumnEnd(multiColumCont) + '<tr><td align="center" style="padding: 30px 4.688%;"><a href="' + f.columLink + '" target="_blank"><img src="' + f.columImage + '" width="auto" border="0" class="full-width" style="display: block;margin:0; max-width:100%;" alt="' + f.columTitle + '"></a></td></tr>';
                        multiColumCont = 0;
                        break;
                    case 'Double-Image':
                        console.log('Double-Image, ' + multiColumCont);
                        codeBody += checkColumnStart(multiColumCont) + '<table border="0" cellpadding="0" cellspacing="0" width="350" class="email-container" style="display: inline-block; vertical-align: top;"><tr><td class="stack-column-center"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td align="center" style="padding: 30px 4.688%;"><a href="' + f.columLink + '" target="_blank"><img src="' + f.columImage + '" width="auto" border="0" class="full-width" style="display: block;margin:0; max-width:100%;" alt="' + f.columTitle + '"></a></td></tr></table></td></tr></table>';
                        multiColumCont += 1;
                        break;
                    case 'Double-column':
                        console.log('Double-column, ' + multiColumCont);
                        codeBody += checkColumnStart(multiColumCont) + '<table border="0" cellpadding="0" cellspacing="0" width="350" class="email-container" style="display: inline-block; vertical-align: top;"><tr><td class="stack-column-center"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td class="full-width-image" style="padding: 20px 10px 0;"><a href="' + f.columLink + '" style="display: block;" target="_blank"><img src="' + f.columImage + '" width="100%" alt="' + f.columTitle + '" border="0" align="center" style="width: 100%; height: auto;"></a></td></tr><tr><td style="padding: 20px 10px 30px;"><a href="' + f.columLink + '" style="display: block; padding: 0 0 10px; text-decoration: none; text-align: left;" target="_blank"><h3 style="margin: 0; color: #000; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 1.5rem; font-weight: 500; mso-height-rule: exactly; line-height: 1.25;">' + f.columTitle + '</h3></a><p style="margin: 0; padding: 0; color: #666; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 1rem; mso-height-rule: exactly; line-height: 1.5; text-align: left;">' + f.columText + '<a href="' + f.columLink + '" style="text-decoration: none;" target="_blank"><span style="display: inline-block; color: #d60c18; text-transform: uppercase; text-decoration: none;">→ MORE</span></a></p></td></tr></table></td></tr></table>';
                        multiColumCont += 1;
                        break;
                    case 'site-title':
                        console.log('site-title, ' + multiColumCont);
                        codeBody += checkColumnEnd(multiColumCont) + '<tr><td style="padding: 0 10px;"><table cellspacing="0" cellpadding="0" border="0" align="center" width="100%" class="email-container"><tr><td style="padding: 5px 0; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 1.125rem; line-height: 1.875rem; border-bottom: 2px solid #000;"><b>' + f.columTitle + '&nbsp;&nbsp;<span style="color: #d60c18;">' + f.columText + '</span></b></td></tr></table></td></tr>';
                        multiColumCont = 0;
                        break;
                    case 'single-column':
                        console.log('single-column, ' + multiColumCont);
                        codeBody += checkColumnEnd(multiColumCont) + '<tr><td style="padding: 20px 10px 0;"><table cellspacing="0" cellpadding="0" border="0" align="center" width="100%" class="email-container"><tr><td class="full-width-image"><a href="' + f.columLink + '" target="_blank"><img src="' + f.columImage + '" width="100%" alt="' + f.columTitle + '" border="0" align="center" style="width: 100%; height: auto;"></a></td></tr><tr><td style="padding: 20px 0 10px;"><a href="' + f.columLink + '" style="display: block; padding: 0 0 10px; text-decoration: none; text-align: left;" target="_blank"><h3 style="margin: 0; color: #000; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 1.5rem; font-weight: 500; mso-height-rule: exactly; line-height: 1.25;">' + f.columTitle + '</h3></a><p style="margin: 0; padding: 0; color: #666; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 1rem; mso-height-rule: exactly; line-height: 1.5; text-align: left;">' + f.columText + '<a href="' + f.columLink + '" style="text-decoration: none;" target="_blank"><span style="display: inline-block; color: #d60c18; text-transform: uppercase; text-decoration: none;">→ MORE</span></a></p></td></tr></table></td></tr>';
                        multiColumCont = 0;
                        break;
                    case 'recommand-book':
                        console.log('recommand-book, ' + multiColumCont);
                        codeBody += checkColumnEnd(multiColumCont) + '<tr><td dir="ltr" align="center" valign="top" width="100%" style="padding: 30px 3.125% 0;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td width="30%" class="stack-column-center"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td dir="ltr" valign="top" style="padding: 0 10px 20px; line-height: 0"><a href="' + f.columLink + '" style="display: block;" target="_blank"><img src="' + f.columImage + '" width="100%" alt="' + f.columText + '" border="0" class="center-on-narrow"></a></td></tr></table></td><td width="70%" class="stack-column-center"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td dir="ltr" valign="top" style="padding: 0 10px;" class="center-on-narrow"><a href="' + f.columLink + '" style="display: block; padding: 0 0 10px; color: #666; line-height: 1.5; word-wrap: break-word; white-space: pre-wrap; text-decoration: none; text-align: left;" target="_blank">' + f.columText + '</a></td></tr></table></td></tr></table></td></tr>';
                        multiColumCont = 0;
                        break;
                    case 'pure-text':
                        console.log('pure-text, ' + multiColumCont);
                        codeBody += checkColumnEnd(multiColumCont) + '<tr><td style="padding: 30px 2.857143% 50px;"><h3 style="margin: 15px 0; font-size: 1.125rem;">' + f.columTitle + '</h3><p style="margin: 0; font-size: 1rem; word-wrap: break-word; white-space: pre-wrap;">' + f.columText + '</p></td></tr>'
                        multiColumCont = 0;
                        break;
                    case 'site-text':
                        console.log('site-text, ' + multiColumCont);
                        codeBody += checkColumnEnd(multiColumCont) + '<tr><td style="padding: 20px 10px 0px;"><h3 style="margin: 0; color: #000; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 1.625rem; font-weight: 500; mso-height-rule: exactly; line-height: 1.25;">' + f.columTitle + '</h3><p style="margin: 10px 0 0; padding: 0; color: #000; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 1.125rem; mso-height-rule: exactly; line-height: 1.5; text-align: left; word-wrap: break-word; white-space: pre-wrap;">' + f.columText + '</p></td></tr>'
                        multiColumCont = 0;
                        break;
                    case 'foot':
                        console.log('foot');
                        break;
                    default:
                        console.log('Nobody sucks!');
                }
            });
        });

        //組裝原始碼第三段，加上表尾及網頁foot的部份
        codeFoot = '</body></html>';

        trackFoot = '<custom name="opencounter" type="tracking"/><div style="display: none; font-size: 1px; line-height: 1px; max-height: 0px; max-width: 0px; opacity: 0; overflow: hidden; mso-hide: all; font-family: sans-serif;">This email was sent by:<b>%%Member_Busname%%</b><br>%%Member_Addr%% %%Member_City%%, %%Member_State%%, %%Member_PostalCode%%, %%Member_Country%%<a href="%%profile_center_url%%" alias="Update Profile">Update Profile</a></div>';

        previewFoot = '</table><table cellspacing="0" cellpadding="0" border="0" align="center" width="700" class="email-container"><tr><td style="padding: 1rem 2.858%; border-top: 3px solid #000;"><table cellspacing="0" cellpadding="0" border="0" width="100%" align="center" class="email-container"><tr><td align="center" style="padding: 0 0 1rem 0; color: #939EA7; font-size: 0.875rem;">此郵件是系統傳送，請勿直接回覆，如有任何問題請<a href="mailto:bill@cw.com.tw" target="_blank" rel="noopener noreferrer" style="color: #6B7780;">聯絡我們</a></td"></tr><tr><td align="center" style="font-size: 0; padding: 5px 0;"><a href="https://www.cw.com.tw/article/5096027?utm_source=email_edm&utm_medium=email" target="_blank" rel="noopener noreferrer" style="display: inline-block; padding: 0 10px; font-size: 0.75rem; color: #000; text-decoration: none; line-height: 1;">關於天下</a><a href="https://account.cwg.tw/register/cw?utm_source=email_edm&utm_medium=email" target="_blank" rel="noopener noreferrer" style="display: inline-block; padding: 0 10px; font-size: 0.75rem; color: #000; text-decoration: none; line-height: 1; border-left: 1px solid #000;">加入會員</a><a href="https://www.cw.com.tw/saleskit/" target="_blank" rel="noopener noreferrer" style="display: inline-block; padding: 0 10px; font-size: 0.75rem; color: #000; text-decoration: none; line-height: 1; border-left: 1px solid #000;">廣告刊登</a><a href="https://www.cw.com.tw/recruit/join1.php" target="_blank" rel="noopener noreferrer" style="display: inline-block; padding: 0 10px; font-size: 0.75rem; color: #000; text-decoration: none; line-height: 1; border-left: 1px solid #000;">人才招募</a></td></tr><tr><td align="center" style="font-size: 0; padding: 5px 0;"><a class="unsubscribe" href="https://api.cwg.tw/web/epaper/cancel-guest?email=%%email%%&epaper=EPAPER21" target="_blank" rel="noopener noreferrer" style="display: inline-block; padding: 0 10px; font-size: 0.75rem; color: #888; text-decoration: none; line-height: 1;">取消訂閱</a><a href="mailto:bill@cw.com.tw" target="_blank" rel="noopener noreferrer" style="display: inline-block; padding: 0 10px; font-size: 0.75rem; color: #888; text-decoration: none; line-height: 1;">讀者服務信箱</a><a href="tel:+886226620332" style="display: inline-block; padding: 0 10px; font-size: 0.75rem; color: #888; text-decoration: none; line-height: 1; border-left: 1px solid #888;">讀者服務專線 (02) 2662-0332</a></td></tr><tr><td align="center" style="font-size: 0.75rem; color: #888;">Copyright © ' + currentYear + ' 天下雜誌 All rights reserved.</td></tr><tr><td align="center" style="padding-top: 20px; font-size: 0"><a href="https://www.facebook.com/cwgroup" target="_blank" rel="noopener noreferrer" style="display: inline-block; margin: 0; padding: 0;" title="天下雜誌 Facebook 粉絲專頁"><img src="https://storage.googleapis.com/cw-com-tw/article/201807/article-5b581cc717ba2.gif" height="100" alt="facebook"></a><a href="https://maac.io/1fqKH" target="_blank" rel="noopener noreferrer" style="display: inline-block; margin: 0; padding: 0;" title="天下雜誌 Line"><img src="https://storage.googleapis.com/cw-com-tw/article/201807/article-5b581cf08d480.gif" height="100" alt="line"></a><a href="https://www.instagram.com/commonwealth_magazine/" target="_blank" rel="noopener noreferrer" style="display: inline-block; margin: 0; padding: 0;" title="天下雜誌 Instagram"><img src="https://storage.googleapis.com/cw-com-tw/article/201807/article-5b5831af69d9f.gif" height="100" alt="line"></a><a href="https://topic.cw.com.tw/cwapp/applink/cwapp.html?source=cwnewsletter-banner-blue" target="_blank" rel="noopener noreferrer" style="display: inline-block; margin: 0; padding: 0;" title="天下雜誌 App"><img src="https://storage.googleapis.com/www-cw-com-tw/article/201810/article-5bc7181566d29.png" height="100" alt="line"></a></td></tr></table>' + trackFoot + '</center></td></tr></table>';

        // dataLayer.push({
        //     "event": "GA-event",
        //     "eventInfo": "編碼完成",
        //     "fileID": key
        // });
        //將原始碼塞入「發信用原始碼」的框框
        $("#sourceCode textarea").val(codeHead+previewHead+codeBody+previewFoot+codeFoot);
        //將原始碼丟入「預覽」框，呈現組裝後的結果
        $("#previewHtml").html(previewHead+codeBody+previewFoot);
        $("#initial").html(codeHead+previewHead+codeBody+previewFoot+codeFoot);
    });
}
function checkColumnStart(multiColumCont) {
    var checkMultiColumn = "<!--checkColumnStart false " + multiColumCont + "-->";
    if (multiColumCont === 0) {
        checkMultiColumn = '<!--checkColumnStart true--><tr><td dir="ltr" align="center" valign="top" width="100%" style="padding: 10px 0; font-size: 0;">';
    }
    return checkMultiColumn;
}
function checkColumnEnd(multiColumCont) {
    var checkMultiColumn = "<!--checkColumnEnd false " + multiColumCont + "-->";
    if (multiColumCont !== 0) {
        checkMultiColumn = "<!--checkColumnEnd true--></td></tr>";
    }
    return checkMultiColumn;
}

$("button#copySource").click(function(){
    $("#sourceCode textarea").select();
    document.execCommand('copy');
});

$("#preview").click(function () {
    var key = $("#dockey").val().split("/d/")[1].split("/")[0],
        worksheet = $("#worksheet").val();
    console.log(key);
    makeSourceCode(key, worksheet);
    $("#previewHtml").show();
    $("#sourceCode").hide(); 
});

$("#previewBlank").click(function () {
    var key = $("#dockey").val().split("/d/")[1].split("/")[0];
    url = initial + "?file=" + key;
    window.open(url, '_blank');
});

$("#source").click( function(){
    var key = $("#dockey").val().split("/d/")[1].split("/")[0],
        worksheet = $("#worksheet").val();
    makeSourceCode(key,worksheet);
    $("#previewHtml").hide();
    $("#sourceCode").show(); 
});

//從網址取得日期參數
function getUrlParameter(sParam) {
    var sPageURL = window.location.search.substring(1);
    var sURLVariables = sPageURL.split('&');
    for (var i = 0; i < sURLVariables.length; i++) {
        var sParameterName = sURLVariables[i].split('=');
        if (sParameterName[0] == sParam) {
            return sParameterName[1];
        }
    }
}

if ( window.location.href.indexOf('initial') > 0 ) {
    var key = getUrlParameter('file'),
        playId = getUrlParameter('playID'),
        data = {
            funcColum: []
        };
    makeSourceCode(key, 1);
    $.getJSON("https://spreadsheets.google.com/feeds/list/" + key + "/1/public/values?alt=json&sq=%E9%A1%AF%E7%A4%BA=y", function(l) {
        //將載入列表轉換為清單的格式
        $.each(l.feed.entry, function(i, f) {
            data.funcColum[i] = {
                columType: f.gsx$欄型.$t,
                columLink: f.gsx$連結.$t,
                columImage: f.gsx$圖片.$t,
                columChannel: f.gsx$小帽.$t,
                columTitle: f.gsx$標題.$t,
                columText: f.gsx$內文.$t
            };
        });
        document.title = data.funcColum[0].columTitle + " - 天下電子報";
		gtag('config', 'UA-1198057-9');
		ga('send', 'pageview');
        if ( window.location.href.indexOf('playID') > 0 ) {
            $('#player').show();
            $('#player').siblings().show();
        }
        if ( window.location.href.indexOf('&email=') > 0 ) {
            var email = window.location.href.split('email=')[1];
            $(window).on('load', function(){
                $('a.unsubscribe').each(function(){
                    var originHref = $(this).attr('href'),
                        unsubscribeHref = originHref.replace('%%email%%', email);
                    $(this).attr('href', unsubscribeHref);
                })
            })
        }
    })
}
