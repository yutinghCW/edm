var width = $(window).width(),
    host = 'https://yutinghcw.github.io/edm',
    initial = host + "/initial/opinion-exp.html",
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
        previewFoot = "",
        codeFoot = "",
        ColumCont = 0;
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
                columTitle: f.gsx$標題.$t,
                columImage: f.gsx$圖片.$t,
                columImageAlt: f.gsx$圖片替代文字.$t,
                columAuthor: f.gsx$作者.$t,
                columText: f.gsx$內文.$t,
                columLink: f.gsx$連結.$t
            };
        });

        //組裝原始碼第一段，版頭及表頭
        codeHead = '<!doctype html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><meta http-equiv="X-UA-Compatible" content="IE=edge"><title>EmailTemplate-Responsive</title><meta content="telephone=no" name="format-detection"><link href="https://fonts.gstatic.com" rel="preconnect" crossorigin><link href="https://fonts.googleapis.com/css?family=Noto+Serif+TC: 300,400,500,700|Noto+Sans+TC: 100,300,400,500,700|Roboto: 100,300,400,500,700&display=swap" rel="stylesheet"><link rel="shortcut icon" href="https://www.cw.com.tw/assets_new/img/favicon.ico" type="image/x-icon" /><style type="text/css">html,body{margin:0 !important;padding:0 !important;height:100% !important;width:100% !important}*{-ms-text-size-adjust:100%;-webkit-text-size-adjust:100%}.ExternalClass{width:100%}div[style*="margin: 16px 0"]{margin:0 !important}table,td{mso-table-lspace:0pt !important;mso-table-rspace:0pt !important}table{border-spacing:0 !important;border-collapse:collapse !important;table-layout:fixed !important;margin:0 auto !important}table table table{table-layout:auto}img{-ms-interpolation-mode:bicubic}.yshortcuts a{border-bottom:none !important}a[x-apple-data-detectors]{color:inherit !important}</style><style type="text/css">.button-td,.button-a{transition:all 100ms ease-in}.button-td:hover,.button-a:hover{background:#555 !important;border-color:#555 !important}@media screen and (max-width: 700px){.email-container{width:100% !important}.fluid,.fluid-centered{max-width:100% !important;height:auto !important;margin-left:auto !important;margin-right:auto !important}.fluid-centered{margin-left:auto !important;margin-right:auto !important}.stack-column,.stack-column-center{display:block !important;width:100% !important;max-width:100% !important;direction:ltr !important}.stack-column-center{text-align:center !important}.center-on-narrow{text-align:center !important;display:block !important;margin-left:auto !important;margin-right:auto !important;float:none !important}table.center-on-narrow{display:inline-block !important}}</style></head><body bgcolor="#ddd" width="100%" style="margin: 0;" yahoo="yahoo">';

        previewHead = '<table bgcolor="#ddd" cellpadding="0" cellspacing="0" border="0" height="100%" width="100%" style="border-collapse:collapse;"><tr><td><center style="width: 100%;"><div style="display:none;font-size:1px;line-height:1px;max-height:0px;max-width:0px;opacity:0;overflow:hidden;mso-hide:all;font-family: sans-serif;">' + data.funcColum[0].columText + '</div><table align="center" width="700" class="email-container"><tr><td style="padding: 3px 15px;"><a class="unsubscribe" href="' + initial + "?file=" + key + '&email=%%email%%" style="color: #888; font-size: 0.875rem; font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif; text-decoration: none;" id="headLink" target="_blank">無法看到完整內容，請點此處</a></td></tr><tr bgcolor="#fff"><td style="padding: 15px 4.688% 10px; text-align: center; line-height: 0;"><a style="display: inline-block;" href="' + data.funcColum[0].columLink + '" target="_blank" rel="noopener noreferrer"><img src="' + data.funcColum[0].columImage + '" style="display: inline-block; height: 80px;" alt="獨立評論@天下"></a></td></tr></table><table cellspacing="0" cellpadding="0" border="0" align="center" bgcolor="#fff" width="700" class="email-container">';

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
                    case 'Editor-column':
                        // console.log('Editor-column');
                        codeBody += '<table cellspacing="0" cellpadding="0" border="0" align="center" bgcolor="#fff" width="700" class="email-container"><tr><td style="padding: 50px 4.688%;"><h1 style="margin-top: 0; margin-bottom: 20px; font-size: 1.375rem; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-weight: 500; line-height: 1;"><span style="padding-left: 10px; border-left: 5px solid #d60c18;">' + f.columTitle + '</span></h1><p style="margin-top: 0; margin-bottom: 0; color: #444; font-size: 1rem; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; line-height: 1.5; white-space: break-spaces;">' + f.columText + '</p><p style="margin-bottom: 0; color: #444; font-size: 1rem; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; text-align: right;">' + f.columAuthor + '</p></td></tr>';
                        break;
                    case 'Site-title':
                        // console.log('Site-title');
                        ColumCont = 0;
                        codeBody += '<tr><td style="padding: 10px 15px; color: #fff; font-size: 1.125rem; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; line-height: 1; background-color: ' + f.columText + ';"><span style="padding-left: 10px; border-left: 5px solid #fff;">' + f.columTitle + '</span></td></tr>';
                        break;
                    case 'Single-column':
                        ColumCont += 1; ColumCont %=2;
                        if (ColumCont){
                            Columbg="#F5F4F2";
                        } else {
                            Columbg="#fff";
                        } //設定該列背景，若非偶數為#FFF
                        // console.log('Single-column');
                        codeBody += '<tr bgcolor="' + Columbg + '"><td style="padding: 30px 20px; text-align: center;"><a href="' + f.columLink + '" style="display: block; padding: 0 0 10px; text-decoration: none; text-align: left;" target="_blank"><h3 style="margin: 0; color: #000; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 1.375rem; font-weight: 500; mso-height-rule: exactly; line-height: 1.35;">' + f.columTitle + '</h3><span style="display: block; color: #888; font-size: 0.9375rem; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif;">' + f.columAuthor + '</span></a><p style="margin: 0; padding: 0 0 10px; color: #444; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 1rem; mso-height-rule: exactly; line-height: 1.5; text-align: left;">' + f.columText + '</p><a href="' + f.columLink + '" style="text-decoration: none;" target="_blank"><span style="display: inline-block; padding: 8px 40px 5px; color: #313160; font-size: 0.9375rem; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-weight: 500; line-height: 1.7; text-decoration: none; background-color: #fff; border: 2px solid #313160;">看全文</span></a></td></tr>';
                        break;
                    case 'Double-column':
                        ColumCont += 1; ColumCont %=2;
                        if (ColumCont){
                            Columbg="#F5F4F2";
                        } else {
                            Columbg="#fff";
                        } //設定該列背景，若非偶數為#FFF
                        // console.log('Double-column');
                        codeBody += '<tr bgcolor="' + Columbg + '"><td dir="ltr" align="center" valign="top" width="100%" style="padding: 30px 10px 0; text-align: center;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td width="45%" valign="top" class="stack-column-center"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td dir="ltr" valign="top" style="padding: 0 10px 30px; line-height: 0"><a href="' + f.columLink + '" style="display: block;" target="_blank"><img src="' + f.columImage + '" width="100%" alt="' + f.columImageAlt + '" border="0" class="center-on-narrow"></a></td></tr></table></td><td width="55%" valign="top" class="stack-column-center"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td dir="ltr" valign="top" style="padding: 0 10px 30px;" class="center-on-narrow"><a href="' + f.columLink + '" style="display: block; padding: 0 0 10px; text-decoration: none; text-align: left;" target="_blank"><h3 style="margin: 0; color: #000; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 1.375rem; font-weight: 500; mso-height-rule: exactly; line-height: 1.35;">' + f.columTitle + '</h3><span style="display: block; color: #888; font-size: 0.9375rem; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif;">' + f.columAuthor + '</span></a><p style="margin: 0; padding: 0 0 10px; color: #444; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 1rem; mso-height-rule: exactly; line-height: 1.5; text-align: left;">' + f.columText + '</p><a href="' + f.columLink + '" style="text-decoration: none;" target="_blank"><span style="display: inline-block; padding: 8px 40px 5px; color: #fff; font-size: 0.9375rem; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-weight: 500; line-height: 1.7; text-decoration: none; background-color: #313160; border: 2px solid #313160; border-radius: 5px;">看全文</span></a></td></tr></table></td></tr></table></td></tr>';
                        break;
                    case 'Single-Image':
                        // console.log('Single-Image');
                        codeBody += '<tr><td align="center"><a href="' + f.columLink + '" target="_blank"><img src="' + f.columImage + '" width="auto" border="0" class="full-width" style="display: block; margin: 0; max-width: 100%;" alt="' + f.columImageAlt + '"></a></td></tr>';
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

        trackFoot = '<custom name="opencounter" type="tracking"/><div style="display: none; font-size: 1px; line-height: 1px; max-height: 0px; max-width: 0px; opacity: 0; overflow: hidden; mso-hide: all; font-family: sans-serif;">This email was sent by:<b>%%Member_Busname%%</b><br>%%Member_Addr%% %%Member_City%%, %%Member_State%%, %%Member_PostalCode%%, %%Member_Country%%<a href="%%profile_center_url%%" alias="Update Profile">Update Profile</a></div>';

        previewFoot = '</table><table cellspacing="0" cellpadding="0" border="0" align="center" bgcolor="#545454" width="700" class="email-container"><tr bgcolor="#545454"><td dir="ltr" align="center" valign="top" width="100%" style="padding: 10px 3.125%;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td width="68" valign="middle" class="stack-column-center"></td><td width="80" valign="middle" class="stack-column-center" style="padding: 10px 0;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td dir="ltr" valign="top" align="center"><a href="https://www.facebook.com/opinion.cw" target="_blank" rel="noopener noreferrer" style="display: inline-block; margin: 0; padding: 0; color: #fff; font-size: 0.75rem; text-decoration: none;" title="天下雜誌 Facebook 粉絲專頁"><img style="display: block; margin-bottom: 5px;" src = host + "/assets/images/facebook-white@2x.png" width="40" alt="facebook">追蹤獨評</a></td></tr></table></td><td width="446" valign="middle" class="stack-column-center"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td dir="ltr" valign="top" style="padding: 0 10px; font-size: 0; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif;" class="center-on-narrow"><a href="https://opinion.cw.com.tw/blog/profile/31/article/18" target="_blank" rel="noopener noreferrer" style="padding: 0 10px 0 0; font-size: 0.75rem; color: #fff;">關於獨立評論@天下</a><a class="unsubscribe" href="https://api.cwg.tw/web/epaper/cancel-guest?email=%%email%%&epaper=EPAPER13" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 0.75rem; color: #fff; border-left: 1px solid #fff;">取消訂閱</a><a href="https://opinion.cw.com.tw/blog/profile/31/article/28" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 0.75rem; color: #fff; border-left: 1px solid #fff;">讀者投書</a><a href="tel:+886226620332" style="padding: 0 0 0 10px; font-size: 0.75rem; color: #fff; border-left: 1px solid #fff; text-decoration: none; cursor: text;">讀者服務專線 (02) 2662-0332</a><p style="padding: 0.5rem 0 0; margin: 0; color: #fff; font-size: 0.75rem; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif;">Copyright © 2021 天下雜誌 All rights reserved.</p></td></tr></table></td><td width="68" valign="middle" class="stack-column-center"></td></tr></table></td></tr></table>' + trackFoot + '</center></td></tr></table>';

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
                columTitle: f.gsx$標題.$t,
                columImage: f.gsx$圖片.$t,
                columImageAlt: f.gsx$圖片替代文字.$t,
                columAuthor: f.gsx$作者.$t,
                columText: f.gsx$內文.$t,
                columLink: f.gsx$連結.$t
            };
        });
        document.title = data.funcColum[0].columTitle + " - 獨立評論電子報";
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
