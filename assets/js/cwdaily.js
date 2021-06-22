var width = $(window).width(),
    host = 'https://yutinghcw.github.io/edm',
    initial = host + "/initial/cwdaily.html",
    initialExp = host + "/initial/cwdaily-exp.html",
    playlist = host + "/initial/playlist-zh.html",
    playlistExp = host + "/initial/playlist-exp.html",
    currentYear = new Date().getFullYear();

//修改目的：提供EDM公版//
//修改1：將電子報版頭資訊改到表單中填入，也讓Speadsheet文件表格單純化//
//修改2：簡化表單，只保留填入Google文件序號欄位，其它都在spreadsheet中編輯//
//修改3：資料表格增加「按鈕文字」欄位//-->
//修改4：調整工具的樣式與SOP，新增多欄版型判斷頭、尾的邏輯//

// 控制組 START --------------------------------
//從參數製作html原始碼
function makeSourceCode(key, worksheet) {
    var codeHead = "",
        previewHead = "",
        adText = "",
        soundLink = "",
        codeFrame = "",
        codeBody = "",
        codeFoot = "",
        previewFoot = "";
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

        if( data.funcColum[1].columType == "AD-text" ){
            adText = '<tr><td style="padding: 0.5rem 2.858%; border-bottom: 1px solid #c9c9c9;"><a href="' + data.funcColum[1].columLink + '" target="_blank" rel="noopener noreferrer" id="adTextLink" style="color:#666; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 0.8125rem; text-decoration: none; line-height: 1.5;">' + data.funcColum[1].columTitle + '</a></td></tr>';
        } else {
            adText = '';
        }

        if( data.funcColum[0].columImage !== "" ){
            soundLink += '<a href="' + initial + "?file=" + key + '&amp;playID=all" title="立刻聽天下" style="float: right; display: block; color: #666; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 0.8125rem; line-height: 1.5; text-decoration: none;"><img style="display: inline-block; vertical-align: middle;" src="https://topic.cw.com.tw/edm/cwdaily/images/headphone@2x.png" width="20" alt="headphone icon"><span style="display: inline-block; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; vertical-align: middle;">&nbsp;&nbsp;立刻聽天下</span></a>';
        }

        if ( window.location.href.indexOf('playID=all') > 0 ) {
            codeFrame = '<tr><td style="padding-top: 20px;"><iframe id="player" width="100%" height="280" scrolling="no" frameborder="no" src="' + playlist + "?file=" + key + '&Daily=' + data.funcColum[0].columImage +'&playID=all' + '" style="display: none;"></iframe></td></tr>';
        }

        //組裝原始碼第一段，版頭及表頭
        codeHead = '<!doctype html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><meta http-equiv="X-UA-Compatible" content="IE=edge"><title>EmailTemplate-Responsive</title><meta content="telephone=no" name="format-detection"><link rel="preconnect" href="https://fonts.gstatic.com"><link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500&family=Roboto:wght@300;400;500&display=swap" rel="stylesheet"><link rel="shortcut icon" href="https://www.cw.com.tw/assets_new/img/favicon.ico" type="image/x-icon" /><style type="text/css">html,body{margin:0 !important;padding:0 !important;height:100% !important;width:100% !important}*{-ms-text-size-adjust:100%;-webkit-text-size-adjust:100%}.ExternalClass{width:100%}div[style*="margin: 16px 0"]{margin:0 !important}table,td{mso-table-lspace:0pt !important;mso-table-rspace:0pt !important}table{border-spacing:0 !important;border-collapse:collapse !important;table-layout:fixed !important;margin:0 auto !important}table table table{table-layout:auto}img{-ms-interpolation-mode:bicubic}.yshortcuts a{border-bottom:none !important}a[x-apple-data-detectors]{color:inherit !important}</style><style type="text/css">.button-td,.button-a{transition:all 100ms ease-in}.button-td:hover,.button-a:hover{background:#555 !important;border-color:#555 !important}@media screen and (max-width: 700px){.email-container{width:100% !important}.fluid,.fluid-centered{max-width:100% !important;height:auto !important;margin-left:auto !important;margin-right:auto !important}.fluid-centered{margin-left:auto !important;margin-right:auto !important}.stack-column,.stack-column-center{display:block !important;width:100% !important;max-width:100% !important;direction:ltr !important}.stack-column-center{text-align:center !important}.center-on-narrow{text-align:center !important;display:block !important;margin-left:auto !important;margin-right:auto !important;float:none !important}table.center-on-narrow{display:inline-block !important}}</style></head><body bgcolor="#fff" width="100%" style="margin: 0;" yahoo="yahoo">';

        previewHead = '<table bgcolor="#fff" cellpadding="0" cellspacing="0" border="0" height="100%" width="100%" style="border-collapse:collapse;"><tr><td><center style="width: 100%;"><div style="display: none; font-size: 1px; line-height: 1px; max-height: 0px; max-width: 0px; opacity: 0; overflow: hidden; mso-hide: all; font-family: sans-serif;">' + data.funcColum[0].columText + '</div><table align="center" width="700" class="email-container"><tr bgcolor="#d50e1a"><td style="padding: 0; text-align: center;"><a style="display: block; line-height: 0;" href="https://www.cw.com.tw/" target="_blank" rel="noopener noreferrer"><img src="https://topic.cw.com.tw/edm/cwdaily/images/1201_logo.jpg" style="display: inline-block; height: 60px;" alt="天下雜誌"></a></td></tr><tr><td style="padding: 1rem 0; text-align: center; border-bottom: 1px solid #c9c9c9;"><p style="margin: 0 0 5px; color:#333; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 0.8125rem; line-height: 1.5;"><span style="color: #999; font-size: 0.8125rem;">' + data.funcColum[0].columLink + '</span><br/>' + data.funcColum[0].columChannel + '</p><p style="margin: 0; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif;"><a class="unsubscribe" href="' + initial + "?file=" + key + '&email=%%email%%" target="_blank" rel="noopener noreferrer" style="color:#999; font-size: 0.8125rem; text-decoration: none;">內容若無法正常顯示，請點我</span></p></td></tr>' + adText + '<tr><td style="padding: 0.5rem 2.858%; border-bottom: 1px solid #c9c9c9;"><a href="https://bit.ly/2zRZJea" style="padding: 0 10px 0 0; color: #666; font-size: 0.8125rem; text-decoration: none;" target="_blank" rel="noopener noreferrer">下載 APP</a><a class="unsubscribe" href="https://api.cwg.tw/web/epaper/cancel-guest?email=%%email%%&epaper=EPAPER19" style="padding: 0 0 0 10px; color: #666; font-size: 0.8125rem; text-decoration: none; border-left: 1px solid #666;" target="_blank" rel="noopener noreferrer">取消訂閱</a>' + soundLink + '</td></tr>' + codeFrame + '</table><table cellspacing="0" cellpadding="0" border="0" align="center" bgcolor="#fff" width="700" class="email-container">';

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
                        codeBody += '<tr><td align="center" style="padding: 1rem 2.858%; border-bottom: 1px solid #c9c9c9;"><a href="' + f.columLink + '" target="_blank"><img src="' + f.columImage + '" height="100" width="auto" border="0" class="full-width" style="display: block;margin:0; max-width:100%;" alt="' + f.columTitle + '"></a></td></tr>';
                        break;
                    case 'Double-column':
                        console.log('Double-column, ' + multiColumCont);
                        codeBody += '<tr><td dir="ltr" align="center" valign="top" width="100%" style="padding: 20px 1.43% 10px; border-bottom: 1px solid #c9c9c9;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td style="padding: 0 10px 10px;"><span style="display: block; color: #D60C18; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 1rem; font-weight: bold;">» ' + f.columChannel + '</span></td></tr></table><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td width="44%" valign="top" class="stack-column-center"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td dir="ltr" valign="top" style="padding: 0 10px 10px; line-height: 0"><a href="' + f.columLink + '" target="_blank" style="display: block;"><img src="' + f.columImage + '" width="100%" alt="alt_text" border="0" class="center-on-narrow"></a></td></tr></table></td><td width="56%" valign="top" class="stack-column-center"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td dir="ltr" valign="top" style="padding: 0 10px 10px;" class="center-on-narrow"><a href="' + f.columLink + '" style="display: block; padding: 0 0 10px; text-align: left; text-decoration: none;" target="_blank"><h3 style="margin: 0; color: #1f4f82; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 1.25rem; font-weight: bold; mso-height-rule: exactly; line-height: 1.5;">' + f.columTitle + '</h3></a><p style="margin: 0; padding: 0; color: #333; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 0.875rem; mso-height-rule: exactly; text-align: left; line-height: 1.5;">' + f.columText + '<a href="' + f.columLink + '" style="color: #1f4f82; text-decoration: none;" target="_blank">...more</a></p></td></tr></table></td></tr></table></td></tr>';
                        break;
                    case 'foot':
                        console.log('foot');
                        break;
                    default:
                        console.log('Nobody sucks!');
                }
            });
            codeBody += '<tr><td width="100%" style="padding: 10px 1.43%;"></td></tr>'
        });

        //組裝原始碼第三段，加上表尾及網頁foot的部份
        codeFoot = '</body></html>';

        trackFoot = '<custom name="opencounter" type="tracking"/><div style="display: none; font-size: 1px; line-height: 1px; max-height: 0px; max-width: 0px; opacity: 0; overflow: hidden; mso-hide: all; font-family: sans-serif;">This email was sent by:<b>%%Member_Busname%%</b><br>%%Member_Addr%% %%Member_City%%, %%Member_State%%, %%Member_PostalCode%%, %%Member_Country%%<a href="%%profile_center_url%%" alias="Update Profile">Update Profile</a></div>';

        previewFoot = '</table><table cellspacing="0" cellpadding="0" border="0" align="center" bgcolor="#545454" width="700" class="email-container"><tr><td style="padding: 1rem 2.858%; color: #939EA7; font-size: 0.875rem; text-align: center; background-color: #fff;">此郵件是系統傳送，請勿直接回覆，如有任何問題請<a href="mailto:bill@cw.com.tw" target="_blank" rel="noopener noreferrer" style="color: #6B7780;">聯絡我們</a></td></tr><tr><td style="padding: 1rem 2.858%;"><table cellspacing="0" cellpadding="0" border="0" width="100%" align="center" class="email-container"><tr><td style="padding-top: 0; font-size: 0; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif;"><a href="https://www.cw.com.tw/cwdaily" target="_blank" rel="noopener noreferrer" style="padding: 0 10px 0 0; font-size: 0.75rem; color: #fff; text-decoration: none;">訂閱每日報</a><a class="unsubscribe" href="https://api.cwg.tw/web/epaper/cancel-guest?email=%%email%%&epaper=EPAPER19" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 0.75rem; color: #fff; text-decoration: none; border-left: 1px solid #fff;">取消訂閱</a><a href="https://www.cw.com.tw/article/articleLogin.action?id=5096027&utm_source=email_edm&utm_medium=email" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 0.75rem; color: #fff; text-decoration: none; border-left: 1px solid #fff;">關於天下</a><a href="https://account.cwg.tw/register/cw?utm_source=email_edm&utm_medium=email" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 0.75rem; color: #fff; text-decoration: none; border-left: 1px solid #fff;">加入會員</a><a href="https://www.cw.com.tw/saleskit/" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 0.75rem; color: #fff; text-decoration: none; border-left: 1px solid #fff;">廣告刊登</a><a href="https://www.cw.com.tw/recruit/join1.php" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 0.75rem; color: #fff; text-decoration: none; border-left: 1px solid #fff;">人才招募</a><a href="mailto:bill@cw.com.tw" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 0.75rem; color: #fff; text-decoration: none; border-left: 1px solid #fff;">讀者服務信箱</a></td></tr><tr><td style="padding: 0.5rem 0; color: #fff; font-size: 0.75rem; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif;">讀者服務專線<a href="tel:+886226620332" style="color: #fff; text-decoration: none;">(02) 2662-0332</a></td></tr><tr><td style="padding: 0.5rem 0; color: #fff; font-size: 0.75rem; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif;">Copyright © ' + currentYear + ' 天下雜誌 All rights reserved.</td></tr><tr><td style="font-size: 0;"><a href="https://www.facebook.com/cwgroup" target="_blank" rel="noopener noreferrer" style="display: inline-block; margin: 0; padding: 0;" title="天下雜誌 Facebook 粉絲專頁"><img src="https://topic.cw.com.tw/edm/cwdaily/images/1202_fb.png" width="80" alt="facebook"></a><a href="https://www.instagram.com/commonwealth_magazine/" target="_blank" rel="noopener noreferrer" style="display: inline-block; margin: 0; padding: 0;" title="天下雜誌 Line"><img src="https://storage.googleapis.com/cw-com-tw/article/201807/article-5b5842e0741c1.png" width="80" alt="line"></a><a href="https://maac.io/1fn12" target="_blank" rel="noopener noreferrer" style="display: inline-block; margin: 0; padding: 0;" title="天下雜誌 Line"><img src="https://topic.cw.com.tw/edm/cwdaily/images/1202_line.png" width="80" alt="line"></a></td></tr></table></td></tr></table>' + trackFoot + '</center></td></tr></table>';

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
// 控制組 END ----------------------------------
// 實驗組 START --------------------------------
function makeSourceCodeExp(key, worksheet) {
    var codeHead = "",
        previewHead = "",
        adText = "",
        soundLink = "",
        codeFrame = "",
        codeBody = "",
        codeFoot = "",
        previewFoot = "";
    var gspreadsheets = "https://spreadsheets.google.com/feeds/list/" + key + "/" + worksheet + "/public/values?alt=json&sq=%E9%A1%AF%E7%A4%BA=y";
    console.log(gspreadsheets);
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

        if( data.funcColum[1].columType == "AD-text" ){
            adText = '<tr><td style="padding: 0.5rem 2.858%; border-bottom: 1px solid #c9c9c9;"><a href="' + data.funcColum[1].columLink + '" target="_blank" rel="noopener noreferrer" id="adTextLink" style="color:#666; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 0.8125rem; text-decoration: none; line-height: 1.5;">' + data.funcColum[1].columTitle + '</a></td></tr>';
        } else {
            adText = '';
        }

        if( data.funcColum[0].columImage !== "" ){
            soundLink += '<a href="' + initialExp + "?file=" + key + '&amp;playID=all" title="立刻聽天下" style="float: right; display: block; color: #666; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 0.8125rem; line-height: 1.5; text-decoration: none;"><img style="display: inline-block; vertical-align: middle;" src="https://topic.cw.com.tw/edm/cwdaily/images/headphone@2x.png" width="20" alt="headphone icon"><span style="display: inline-block; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; vertical-align: middle;">&nbsp;&nbsp;立刻聽天下</span></a>';
        }

        if ( window.location.href.indexOf('playID=all') > 0 ) {
            codeFrame = '<tr><td style="padding-top: 20px;"><iframe id="player" width="100%" height="250" scrolling="no" frameborder="no" src="' + playlistExp + "?file=" + key + '&Daily=' + data.funcColum[0].columImage +'&playID=all' + '" style="display: none;"></iframe></td></tr>';
        }

        //組裝原始碼第一段，版頭及表頭
        previewHead = '<!DOCTYPE html><html><head><meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0"/><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/><style type="text/css">ReadMsgBody{width: 100%;}.ExternalClass{width: 100%;}.ExternalClass, .ExternalClass p, .ExternalClass span, .ExternalClass font, .ExternalClass td, .ExternalClass div{line-height: 100%;}body{-webkit-text-size-adjust:100%; -ms-text-size-adjust:100%;margin:0 !important;}p{margin: 1em 0;}table td{border-collapse: collapse;}img{outline:0;}a img{border:none;}@-ms-viewport{width: device-width;}</style><style type="text/css">@media only screen and (max-width: 480px){.container{width: 100% !important;}.footer{width:auto !important; margin-left:0;}.mobile-hidden{display:none !important;}.logo{display:block !important; padding:0 !important;}img{max-width:100% !important; height:auto !important; max-height:auto !important;}.header img{max-width:100% !important;height:auto !important; max-height:auto !important;}.photo img{width:100% !important; max-width:100% !important; height:auto !important;}.drop{display:block !important; width: 100% !important; float:left; clear:both;}.footerlogo{display:block !important; width: 100% !important; padding-top:15px; float:left; clear:both;}.nav4, .nav5, .nav6{display: none !important;}.tableBlock{width:100% !important;}.responsive-td{width:100% !important; display:block !important; padding:0 !important;}.fluid, .fluid-centered{width: 100% !important; max-width: 100% !important; height: auto !important; margin-left: auto !important; margin-right: auto !important;}.fluid-centered{margin-left: auto !important; margin-right: auto !important;}/* MOBILE GLOBAL STYLES - DO NOT CHANGE */ body{padding: 0px !important; font-size: 1rem !important; line-height: 150% !important;}h1{font-size: 1.375rem !important; line-height: normal !important;}h2{font-size: 1.25rem !important; line-height: normal !important;}h3{font-size: 1.125rem !important; line-height: normal !important;}.buttonstyles{font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 1rem !important; color: #FFFFFF !important; padding: 10px !important;}/* END OF MOBILE GLOBAL STYLES - DO NOT CHANGE */}@media only screen and (max-width: 640px){.container{width:100% !important;}.mobile-hidden{display:none !important;}.logo{display:block !important; padding:0 !important;}.photo img{width:100% !important; height:auto !important;}.nav5, .nav6{display: none !important;}.fluid, .fluid-centered{width: 100% !important; max-width: 100% !important; height: auto !important; margin-left: auto !important; margin-right: auto !important;}.fluid-centered{margin-left: auto !important; margin-right: auto !important;}}</style><!--[if mso]><style type="text/css">/* Begin Outlook Font Fix */ body, table, td{font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 1rem; color:#000000; line-height:1;}/* End Outlook Font Fix */</style><![endif]--></head><body bgcolor="#ffffff" text="#000000" style="background-color: #ffffff; color: #000000; padding: 0px; -webkit-text-size-adjust:none; font-size: 1rem; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif;"><div style="font-size:0; line-height:0;"><custom name="opencounter" type="tracking"><custom name="usermatch" type="tracking"/></div>';

        codeHead = '<table width="100%" border="0" cellpadding="0" cellspacing="0" align="center"><tr><td align="center"><table cellspacing="0" cellpadding="0" border="0" width="600" class="container" align="center"><tr><td><table class="tb_properties border_style" style="background-color:#FFFFFF;" cellspacing="0" cellpadding="0" bgcolor="#ffffff" width="100%"><tr><td align="center" valign="top"><table align="left" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td class="content_padding" style=""><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td align="center" class="header" valign="top"><table align="left" border="0" cellpadding="0" cellspacing="0" width="100%"><tbody><tr><td align="left" valign="top"><table cellspacing="0" cellpadding="0" style="width:100%"><tbody><tr><td class="responsive-td" valign="top" style="width: 100%;">' + '<table cellpadding="0" cellspacing="0" width="100%" role="presentation" style="min-width: 100%; " class="stylingblock-content-wrapper"><tr><td class="stylingblock-content-wrapper camarker-inner"><table cellspacing="0" cellpadding="0" border="0" width="700" class="container" align="center"><tr><td><table cellspacing="0" cellpadding="0" border="0" align="center" width="700" class="email-container"><tr bgcolor="#d50e1a"><td align="center"><a style="display: block; font-size: 0; text-align: center;" href="https://www.cw.com.tw" conversion="false"><img src="https://image.cw.com.tw/lib/fe33117171640474741074/m/1/517610e8-7e7e-40b0-9161-b4b5ff8fb354.jpg" height="60" width="109" style="display: inline-block; height: 60px; width: 109px; padding: 0px;" alt="天下雜誌"></a></td></tr><tr><td style="padding: 1rem 0; text-align: center; border-bottom: 1px solid #c9c9c9;"><p style="margin: 0 0 5px; color:#333; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 0.8125rem; line-height: 1.5;"><span style="color: #999; font-size: 0.8125rem;">' + data.funcColum[0].columLink + '</span><br/>' + data.funcColum[0].columChannel + '</p><p style="margin: 0; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif;"><a class="unsubscribe" href="' + initialExp + "?file=" + key + '&email=%%email%%" target="_blank" rel="noopener noreferrer" style="color:#999; font-size: 0.8125rem; text-decoration: none;">內容若無法正常顯示，請點我</span></p></td></tr>' + adText + '<tr><td style="padding: 0.5rem 2.858%; border-bottom: 1px solid #c9c9c9;"><a href="https://bit.ly/2zRZJea" style="padding: 0 10px 0 0; color: #666; font-size: 0.8125rem; text-decoration: none;" target="_blank" rel="noopener noreferrer">下載 APP</a><a class="unsubscribe" href="https://api.cwg.tw/web/epaper/cancel-guest?email=%%email%%&epaper=EPAPER19" style="padding: 0 0 0 10px; color: #666; font-size: 0.8125rem; text-decoration: none; border-left: 1px solid #666;" target="_blank" rel="noopener noreferrer">取消訂閱</a>' + soundLink + '</td></tr>' + codeFrame + '</table></td></tr></table></td></tr></table>';

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
                        codeBody += '<table cellpadding="0" cellspacing="0" width="100%" role="presentation" style="min-width: 100%; " class="stylingblock-content-wrapper"><tr><td class="stylingblock-content-wrapper camarker-inner"><table cellspacing="0" cellpadding="0" border="0" width="700" class="container" align="center"><tr><td class="slot-styling camarker-inner"><table cellpadding="0" cellspacing="0" width="100%" role="presentation" style="min-width: 100%; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 1.125rem; line-height: 1.875rem; border-bottom: 2px solid #000;" class="stylingblock-content-wrapper"><tr><td class="stylingblock-content-wrapper camarker-inner"><table width="100%" cellspacing="0" cellpadding="0" role="presentation"><tr><td align="center" style="padding: 1rem 2.858%; border-bottom: 1px solid #c9c9c9;"><a href="' + f.columLink + '" target="_blank"><img src="' + f.columImage + '" height="100" width="auto" border="0" class="full-width" style="display: block;margin:0; max-width:100%;" alt="' + f.columTitle + '"></a></td></tr></table></td></tr></table></td></tr></table></td></tr></table>';
                        break;
                    case 'Section-title':
                        codeBody += '<table cellpadding="0" cellspacing="0" width="100%" role="presentation" style="min-width: 100%; " class="stylingblock-content-wrapper"><tr><td class="stylingblock-content-wrapper camarker-inner"><table cellspacing="0" cellpadding="0" border="0" width="700" class="container" align="center"><tr><td class="slot-styling camarker-inner"><table cellpadding="0" cellspacing="0" width="100%" role="presentation" style="min-width: 100%; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 1.125rem; line-height: 1.875rem; border-bottom: 2px solid #000;" class="stylingblock-content-wrapper"><tr><td class="stylingblock-content-wrapper camarker-inner"><table width="100%" cellspacing="0" cellpadding="0" role="presentation"><tr><td style="padding: 20px 10px 5px;"><b>編輯選文&nbsp;&nbsp;<span style="color: #d60c18;">FEATURES</span></b></td></tr></table></td></tr></table></td></tr></table></td></tr></table>';
                        break;
                    case 'Double-column':
                        console.log('Double-column, ' + multiColumCont);
                        codeBody += '<table cellpadding="0" cellspacing="0" width="100%" role="presentation" style="min-width: 100%; " class="stylingblock-content-wrapper"><tr><td class="stylingblock-content-wrapper camarker-inner"><table cellspacing="0" cellpadding="0" border="0" width="700" class="container" align="center"><tr><td><table style="background-color : #FFFFFF; border : 0px solid transparent;font-size : 16px; font-family : Arial, helvetica, sans-serif; line-height : 150%; color : #171717; " class="tb_properties border_style" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF" width="100%"><tr><td align="center" valign="top"><table align="left" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td class="content_padding" style="border:0px;padding : 0px; "><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td align="left" class="" valign="top"><table cellpadding="0" cellspacing="0" width="100%" role="presentation" style="min-width: 100%; " class="slot-styling"><tr><td style="padding: 20px 0 10px; border-bottom: 1px solid #c9c9c9" class="slot-styling camarker-inner"><table cellpadding="0" cellspacing="0" width="100%" role="presentation" style="text-align: left; min-width: 100%; " class="stylingblock-content-wrapper"><tr><td style="padding: 0 10px;"><span style="display: block; color: #D60C18; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 1rem; font-weight: bold;">» ' + f.columChannel + '</span></td></tr><tr><td style="padding: 0;" class="stylingblock-content-wrapper camarker-inner" align="left"><table cellspacing="0" cellpadding="0" style="width: 100%;"><tr><td><table cellspacing="0" cellpadding="0" style="width: 100%;"><tr><td class="responsive-td" valign="middle" style="width: 44%; padding-left: 10px; padding-right: 10px;"><table cellpadding="0" cellspacing="0" width="100%" role="presentation" style="min-width: 100%; " class="stylingblock-content-wrapper"><tr><td style="padding-top: 10px; padding-bottom: 10px;" class="stylingblock-content-wrapper camarker-inner"><a href="' + f.columLink + '" target="_blank" rel="noopener noreferrer" style="display: block;"><img src="' + f.columImage + '" height="auto" width="100%" style="height:auto;width:100%;display: block;"></a></td></tr></table></td><td class="responsive-td" valign="middle" style="width: 56%; padding-left: 10px; padding-right: 10px;"><table cellpadding="0" cellspacing="0" width="100%" role="presentation" style="text-align: left; min-width: 100%; " class="stylingblock-content-wrapper"><tr><td class="stylingblock-content-wrapper camarker-inner" align="left"><a href="' + f.columLink + '" rel="noopener noreferrer" style="display: block; padding: 10px 0; text-align: left; text-decoration: none; color: #1f4f82; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 1.25rem; font-weight: bold; line-height: 1.5;" target="_blank">' + f.columTitle + '</a><p style="margin: 0; padding: 0 0 10px; color: #333; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif; font-size: 0.875rem; text-align: left; line-height: 1.5;">' + f.columText + '<a href="' + f.columLink + '" style="color: #1f4f82; text-decoration: none;" target="_blank">...more</a></p></td></tr></table></td></tr></table></td></tr></table></td></tr></table></td></tr></table></td></tr></table></td></tr></table></td></tr></table></td></tr></table></td></tr></table>';
                        break;
                    case 'foot':
                        console.log('foot');
                        break;
                    default:
                        console.log('Nobody sucks!');
                }
            });
            codeBody += '<table cellpadding="0" cellspacing="0" width="100%" role="presentation" style="min-width: 100%; " class="stylingblock-content-wrapper"><tr><td class="stylingblock-content-wrapper camarker-inner"><tr><td width="100%" style="padding: 10px 1.43%;"></td></tr></td></tr></table>'
        });

        //組裝原始碼第三段，加上表尾及網頁foot的部份
        codeFoot = '<table cellpadding="0" cellspacing="0" width="100%" role="presentation" style="min-width: 100%; " class="stylingblock-content-wrapper"><tr><td class="stylingblock-content-wrapper camarker-inner"><table cellspacing="0" cellpadding="0" border="0" align="center" bgcolor="#545454" width="700" class="email-container"><tr><td style="padding: 1rem 2.858%; color: #939EA7; font-size: 0.875rem; text-align: center; background-color: #fff;">此郵件是系統傳送，請勿直接回覆，如有任何問題請<a href="mailto:bill@cw.com.tw" target="_blank" rel="noopener noreferrer" style="color: #6B7780;">聯絡我們</a></td></tr><tr><td style="padding: 1rem 2.858%;"><table cellspacing="0" cellpadding="0" border="0" width="100%" align="center" class="email-container"><tr><td style="padding-top: 0; font-size: 0; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif;"><a href="https://www.cw.com.tw/cwdaily" target="_blank" rel="noopener noreferrer" style="padding: 0 10px 0 0; font-size: 0.75rem; color: #fff; text-decoration: none;">訂閱每日報</a><a class="unsubscribe" href="https://api.cwg.tw/web/epaper/cancel-guest?email=%%email%%&epaper=EPAPER19" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 0.75rem; color: #fff; text-decoration: none; border-left: 1px solid #fff;">取消訂閱</a><a href="https://www.cw.com.tw/article/articleLogin.action?id=5096027&utm_source=email_edm&utm_medium=email" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 0.75rem; color: #fff; text-decoration: none; border-left: 1px solid #fff;">關於天下</a><a href="https://account.cwg.tw/register/cw?utm_source=email_edm&utm_medium=email" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 0.75rem; color: #fff; text-decoration: none; border-left: 1px solid #fff;">加入會員</a><a href="https://www.cw.com.tw/saleskit/" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 0.75rem; color: #fff; text-decoration: none; border-left: 1px solid #fff;">廣告刊登</a><a href="https://www.cw.com.tw/recruit/join1.php" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 0.75rem; color: #fff; text-decoration: none; border-left: 1px solid #fff;">人才招募</a><a href="mailto:bill@cw.com.tw" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 0.75rem; color: #fff; text-decoration: none; border-left: 1px solid #fff;">讀者服務信箱</a></td></tr><tr><td style="padding: 0.5rem 0; color: #fff; font-size: 0.75rem; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif;">讀者服務專線<a href="tel:+886226620332" style="color: #fff; text-decoration: none;">(02) 2662-0332</a></td></tr><tr><td style="padding: 0.5rem 0; color: #fff; font-size: 0.75rem; font-family: roboto, noto sans tc, 思源黑體 tc, 思源黑體 tw, 思源黑體, 微軟正黑體, 繁黑體, microsoft jhenghei, arial, sans-serif;">Copyright © 2021 天下雜誌 All rights reserved.</td></tr><tr><td style="font-size: 0;"><a href="https://www.facebook.com/cwgroup" target="_blank" rel="noopener noreferrer" style="display: inline-block; margin: 0; padding: 0;" title="天下雜誌 Facebook 粉絲專頁"><img src="https://topic.cw.com.tw/edm/cwdaily/images/1202_fb.png" width="80" alt="facebook"></a><a href="https://www.instagram.com/commonwealth_magazine/" target="_blank" rel="noopener noreferrer" style="display: inline-block; margin: 0; padding: 0;" title="天下雜誌 Line"><img src="https://storage.googleapis.com/cw-com-tw/article/201807/article-5b5842e0741c1.png" width="80" alt="line"></a><a href="https://maac.io/1fn12" target="_blank" rel="noopener noreferrer" style="display: inline-block; margin: 0; padding: 0;" title="天下雜誌 Line"><img src="https://topic.cw.com.tw/edm/cwdaily/images/1202_line.png" width="80" alt="line"></a></td></tr></table></td></tr></table></td></tr></table>';

        trackFoot = '<custom name="opencounter" type="tracking"/><div style="display: none; font-size: 1px; line-height: 1px; max-height: 0px; max-width: 0px; opacity: 0; overflow: hidden; mso-hide: all; font-family: sans-serif;">This email was sent by:<b>%%Member_Busname%%</b><br>%%Member_Addr%% %%Member_City%%, %%Member_State%%, %%Member_PostalCode%%, %%Member_Country%%<a href="%%profile_center_url%%" alias="Update Profile">Update Profile</a></div>';

        previewFoot1 = ' </td></tr></tbody></table></td></tr></tbody></table></td></tr></table></td></tr></table></td></tr></table></td></tr></table></td></tr></table>';
        previewFoot2 = ' </body></html>';

        // dataLayer.push({
        //     "event": "GA-event",
        //     "eventInfo": "編碼完成",
        //     "fileID": key
        // });
        //將原始碼塞入「發信用原始碼」的框框
        $("#sourceCodeExp #textareaA").val(codeHead);
        $("#sourceCodeExp #textareaB").val(codeBody);
        $("#sourceCodeExp #textareaC").val(codeFoot+trackFoot);
        //將原始碼丟入「預覽」框，呈現組裝後的結果
        $("#previewHtmlExp").html(codeHead+codeBody+codeFoot+previewFoot1);
        $("#initialExp").html(previewHead+codeHead+codeBody+codeFoot+previewFoot1+previewFoot2);
    });
}
// 實驗組 END ----------------------------------
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

$('button.copySource').each(function(){
    $(this).click(function(){
        $(this).next('textarea').select();
        document.execCommand('copy');
    });
});

$("#preview").click(function () {
    var key = $("#dockey").val().split("/d/")[1].split("/")[0],
        worksheet = $("#worksheet").val();
    console.log(key);
    makeSourceCode(key, worksheet);
    $("#sourceCode, #previewHtmlExp, #sourceCodeExp").hide(); 
    $("#previewHtml").show();
});

$("#previewExp").click(function () {
    var key = $("#dockey").val().split("/d/")[1].split("/")[0],
        worksheet = $("#worksheet").val();
    console.log(key);
    makeSourceCodeExp(key, worksheet);
    $("#sourceCode, #previewHtml, #sourceCodeExp").hide(); 
    $("#previewHtmlExp").show();
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
    $("#previewHtml, #previewHtmlExp, #sourceCodeExp").hide();
    $("#sourceCode").show(); 
});

$("#sourceExp").click( function(){
    var key = $("#dockey").val().split("/d/")[1].split("/")[0],
        worksheet = $("#worksheet").val();
        makeSourceCodeExp(key,worksheet);
    $("#previewHtml, #previewHtmlExp, #sourceCode").hide();
    $("#sourceCodeExp").show(); 
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
    makeSourceCodeExp(key, 1);
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
        document.title = data.funcColum[0].columTitle + " - 天下每日報";
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

