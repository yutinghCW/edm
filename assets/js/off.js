var width = $(window).width(),
    initial = "https://yutinghcw.github.io/edm/initial/off.html",
    currentYear = new Date().getFullYear();

//修改目的：提供EDM公版//
//修改1：將電子報版頭資訊改到表單中填入，也讓Speadsheet文件表格單純化//
//修改2：簡化表單，只保留填入Google文件序號欄位，其它都在spreadsheet中編輯//
//修改3：資料表格增加「按鈕文字」欄位//-->
//修改4：調整工具的樣式與SOP，新增多欄版型判斷頭、尾的邏輯//

//從參數製作html原始碼
function makeSourceCode(key, worksheet) {
    var codeHead = "",
        codeBody = "",
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

        var mainColor = data.funcColum[0].columChannel,
            textColor = data.funcColum[0].columImage;

        //組裝原始碼第一段，版頭及表頭
        codeHead = '<!doctype html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><meta http-equiv="X-UA-Compatible" content="IE=edge"><title>EmailTemplate-Responsive</title><meta content="telephone=no" name="format-detection"><link rel="preconnect" href="https://fonts.gstatic.com"><link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500&family=Roboto:wght@300;400;500&display=swap" rel="stylesheet"><link rel="shortcut icon" href="https://www.cw.com.tw/assets_new/img/favicon.ico" type="image/x-icon" /><style type="text/css">html,body{margin:0 !important;padding:0 !important;height:100% !important;width:100% !important}*{-ms-text-size-adjust:100%;-webkit-text-size-adjust:100%}.ExternalClass{width:100%}div[style*="margin: 16px 0"]{margin:0 !important}table,td{mso-table-lspace:0pt !important;mso-table-rspace:0pt !important}table{border-spacing:0 !important;border-collapse:collapse !important;table-layout:fixed !important;margin:0 auto !important}table table table{table-layout:auto}img{-ms-interpolation-mode:bicubic}.yshortcuts a{border-bottom:none !important}a[x-apple-data-detectors]{color:inherit !important}</style><style type="text/css">.button-td,.button-a{transition:all 100ms ease-in}.button-td:hover,.button-a:hover{background:#555 !important;border-color:#555 !important}@media screen and (max-width: 700px){.email-container{width:100% !important}.fluid,.fluid-centered{max-width:100% !important;height:auto !important;margin-left:auto !important;margin-right:auto !important}.fluid-centered{margin-left:auto !important;margin-right:auto !important}.stack-column,.stack-column-center{display:block !important;width:100% !important;max-width:100% !important;direction:ltr !important}.stack-column-center{text-align:center !important}.center-on-narrow{text-align:center !important;display:block !important;margin-left:auto !important;margin-right:auto !important;float:none !important}table.center-on-narrow{display:inline-block !important}}</style></head><body bgcolor="#efefef" width="100%" style="margin: 0;" yahoo="yahoo">';

        previewHead = '<table bgcolor="#efefef" cellpadding="0" cellspacing="0" border="0" height="100%" width="100%" style="border-collapse: collapse;"><tr><td><center style="width: 100%;"><div style="display:none;font-size:1px;line-height:1px;max-height:0px;max-width:0px;opacity:0;overflow:hidden;mso-hide:all;font-family: sans-serif;">' + data.funcColum[0].columText + '</div><table align="center" width="640" class="email-container"><tr bgcolor="#efefef"><td style="padding:10px 15px; text-align: right;"><a class="unsubscribe" href="' + initial + "?file=" + key + '&email=%%email%%" style="color: #48545D; font-size: 0.875rem; font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif; text-decoration: none;" id="headLink" target="_blank">無法看到完整內容，請點此處</a></td></tr><tr bgcolor="#fff"><td style="padding: 30px 4.688% 20px; font-size: 0; text-align: center;"><a href="https://www.cw.com.tw/" style="display: inline-block; padding: 0 5px 0 0; line-height: 0; vertical-align: middle;" target="_blank" rel="noopener noreferrer"><img src="https://topic.cw.com.tw/edm/cw_transformers/img/cw_logo.jpg" height="33" alt="天下雜誌"></a><a href="https://www.cw.com.tw/transformers" style="display: inline-block; padding: 0 0 0 5px; line-height: 0; vertical-align: middle;" target="_blank" rel="noopener noreferrer"><img src="https://topic.cw.com.tw/edm/cw_off/img/off_logo.jpg" height="40" alt="Off學"></a><div style="margin-top: 10px; line-height: 1.5;"><div style="font-size: 1rem; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif;"><strong style="font-weight: 500;">' + data.funcColum[0].columTitle + '</strong></div><div style="font-size: 0.875rem; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif;">' + data.funcColum[0].columLink + '</div></div></td></tr><tr><td style="padding: 0 4.688%; line-height: 0; border-top: 1px solid #c9c9c9;"></td></tr></table><table cellspacing="0" cellpadding="0" border="0" align="center" bgcolor="#fff" width="640" class="email-container">';

        //組裝原始碼第二段，填入spreadsheet的內容部份
        var multiColumCont = 0;
        $(function() {
            $.each(data.funcColum, function(i, f) {
                if (i == 0) {
                    return true;
                }
                switch (f.columType) {
                    case 'head':
                        // console.log('head');
                        break;
                    case 'main-title':
                        // console.log('main-title');
                        codeBody += '<tr><td style="padding: 50px 4.688% 0; text-align: center;"><h1 style="margin: 0; font-size: 2rem; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-weight: 300; line-height: 1.25;"><strong style="font-weight: 500;">' + f.columTitle + '</strong><br/>' + f.columChannel + '</h1><p style="margin: 40px 0 0; color: #888; font-size: 1rem; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; text-align: left; line-height: 1.5; white-space: break-spaces;">' + f.columText + '</p><p style="margin: 20px 0; font-size: 1rem; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; text-align: right;">' + f.columImage + '</p></td></tr>';
                        break;
                    case 'single-column':
                        // console.log('single-column');
                        codeBody += '<tr><td style="padding: 20px 15px 0;"><table cellspacing="0" cellpadding="0" border="0" align="center" width="100%" class="email-container"><tr><td class="full-width-image"><a href="' + f.columLink + '" style="display: block;" target="_blank"><img src="' + f.columImage + '" width="640" alt="' + f.columTitle + '" border="0" align="center" style="width: 100%; max-width: 640px; height: auto;"></a></td></tr><tr><td style="padding: 20px 0 0;"><a href="' + f.columLink + '" style="display: block; padding: 0 0 10px; text-decoration: none; text-align: left;" target="_blank"><h3 style="margin: 0; color: #000; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 1.5rem; font-weight: 400; mso-height-rule: exactly; line-height: 1.5;">' + f.columTitle + '</h3></a><p style="margin: 0; padding: 0 0 20px; color: #939EA7; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 1rem; mso-height-rule: exactly; line-height: 1.5; text-align: left;">' + f.columText + '&nbsp;&nbsp;<a href="' + f.columLink + '" style="text-decoration: none;" target="_blank"><span style="display: inline-block; padding: 2px 10px 0; color: #000; font-size: 0.875rem; text-decoration: none; border: 1px solid #D7B788;">看更多</span></a></p></td></tr></table></td></tr>';
                        break;
                    case 'Double-column':
                        // console.log('Double-column');
                        codeBody += '<tr><td dir="ltr" align="center" valign="top" width="100%" style="padding: 20px 5px;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td width="45%" valign="middle" class="stack-column-center"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td dir="ltr" valign="top" style="padding: 0 10px; line-height: 0"><a href="' + f.columLink + '" style="display: block;" target="_blank"><img src="' + f.columImage + '" width="100%" alt="' + f.columTitle + '" border="0" class="center-on-narrow"></a></td></tr></table></td><td width="55%" valign="middle" class="stack-column-center"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td dir="ltr" valign="top" style="padding: 10px;" class="center-on-narrow"><a href="' + f.columLink + '" style="display: block; padding: 0 0 10px; text-align: left; text-decoration: none;" target="_blank"><h3 style="margin: 0; color: #000; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 1.5rem; font-weight: 400; mso-height-rule: exactly; line-height: 1.5;">' + f.columTitle + '</h3></a><p style="margin: 0; padding: 0; color: #939EA7; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 1rem; mso-height-rule: exactly; text-align: left; line-height: 1.5;">' + f.columText + '&nbsp;&nbsp;<a href="' + f.columLink + '" style="text-decoration: none;" target="_blank"><span style="display: inline-block; padding: 2px 10px 0; color: #000; font-size: 0.875rem; text-decoration: none; border: 1px solid #D7B788;">看更多</span></a></p></td></tr></table></td></tr></table></td></tr>';
                        break;
                    case 'subscription-block':
                            // console.log('Single-Image');
                        codeBody += '<tr><td style="padding: 20px 15px 0;"><table cellspacing="0" cellpadding="0" border="0" align="center" width="100%" class="email-container"><tr><td><h3 style="margin-top: 20px; margin-bottom: 10px; color: #000; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 1.5rem; font-weight: 400; mso-height-rule: exactly; line-height: 1.5; text-align: center;">' + f.columTitle + '</h3><p style="margin-top: 0; margin-bottom: 30px; color: #888; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif;">' + f.columText + '</p></td></tr><tr><td class="full-width-image"><a href="' + f.columLink + '" style="display: block;" target="_blank"><img src="' + f.columImage + '" width="640" alt="alt_text" border="0" align="center" style="width: 100%; max-width: 640px; height: auto;"></a></td></tr></table></td></tr>';
                        break;
                    case 'foot':
                        // console.log('foot');
                        break;
                    default:
                        // console.log('Nobody sucks!');
                }
            });
        });

        //組裝原始碼第三段，加上表尾及網頁foot的部份
        codeFoot = '</body></html>';

        previewFoot = '</table><table cellspacing="0" cellpadding="0" border="0" align="center" bgcolor="#fff" width="640" class="email-container"><tr><td style="padding: 30px 4.688%; text-align: center;"><table cellspacing="0" cellpadding="0" border="0" width="100%" align="center" class="email-container"><tr><td align="center">FOLLOW　US</td></tr><tr><td style="padding-top: 10px; font-size: 0;"><a href="https://www.facebook.com/cwgroup" target="_blank" rel="noopener noreferrer" style="display: inline-block; margin: 0 5px;" title="天下雜誌 Facebook 粉絲專頁"><img src="https://topic.cw.com.tw/edm/cw_transformers/img/facebook.png" height="45" alt="facebook"></a><a href="https://maac.io/1fn12" target="_blank" rel="noopener noreferrer" style="display: inline-block; margin: 0 5px;" title="天下雜誌 Line"><img src="https://topic.cw.com.tw/edm/cw_transformers/img/line.png" height="45" alt="line"></a><a href="https://www.instagram.com/commonwealth_magazine/" target="_blank" rel="noopener noreferrer" style="display: inline-block; margin: 0 5px;" title="天下雜誌 Instagram"><img src="https://topic.cw.com.tw/edm/cw_transformers/img/instagram.png" height="45" alt="instagram"></a></td></tr><tr><td align="center" style="padding-top: 20px; font-size: 0; font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif;"><a href="https://www.cw.com.tw/article/articleLogin.action?id=5096027&amp;utm_source=email_edm&amp;utm_medium=email" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 1rem; color: #888; text-decoration: none;">關於天下</a><a href="https://account.cwg.tw/register/cw?utm_source=email_edm&amp;utm_medium=email" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 1rem; color: #888; text-decoration: none; border-left: 1px solid #939EA7;">加入會員</a><a href="https://www.cw.com.tw/saleskit/" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 1rem; color: #888; text-decoration: none; border-left: 1px solid #939EA7;">廣告刊登</a><a href="https://www.cw.com.tw/recruit/join1.php" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 1rem; color: #888; text-decoration: none; border-left: 1px solid #939EA7;">人才招募</a></td></tr><tr><td align="center" style="padding-top: 5px; font-size: 0; font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif;"><a class="unsubscribe" href="http://api.cwg.tw/web/epaper/cancel-guest?email=%%email%%&epaper=EPAPER26" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 1rem; color: #888; text-decoration: none;">取消訂閱</a><a href="mailto:bill@cw.com.tw" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 1rem; color: #888; text-decoration: none; border-left: 1px solid #939EA7;">讀者服務信箱</a><a href="tel:0226620332" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 1rem; color: #888; text-decoration: none; border-left: 1px solid #939EA7;">讀者服務專線 (02) 2662-0332</a></td></tr><tr><td align="center" style="padding-top: 10px; font-size: 1rem; color: #888; font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif;">Copyright © ' + currentYear + ' 天下雜誌 All rights reserved.</td></tr></table></td></tr></table><custom name="opencounter" type="tracking"/><div style="display: none; font-size: 1px; line-height: 1px; max-height: 0px; max-width: 0px; opacity: 0; overflow: hidden; mso-hide: all; font-family: sans-serif;">This email was sent by:<b>%%Member_Busname%%</b><br>%%Member_Addr%% %%Member_City%%, %%Member_State%%, %%Member_PostalCode%%, %%Member_Country%%<a href="%%profile_center_url%%" alias="Update Profile">Update Profile</a></div></center></td></tr></table>';

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
    })
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
