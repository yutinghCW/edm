var width = $(window).width(),
    initial = "https://yutinghcw.github.io/edm/initial/cweng.html",
    playlist = "https://yutinghcw.github.io/edm/initial/playlist-en.html",
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
    var gspreadsheets = "https://spreadsheets.google.com/feeds/list/" + key + "/" + worksheet + "/public/values?alt=json&sq=display=y";

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
        codeHead = '<!doctype html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><meta http-equiv="X-UA-Compatible" content="IE=edge"><title>EmailTemplate-Responsive</title><meta content="telephone=no" name="format-detection"><link rel="preconnect" href="https://fonts.gstatic.com"><link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500&family=Roboto:wght@300;400;500&display=swap" rel="stylesheet"><link rel="shortcut icon" href="https://www.cw.com.tw/assets_new/img/favicon.ico" type="image/x-icon" /><style type="text/css">html,body{margin:0 !important;padding:0 !important;height:100% !important;width:100% !important}*{-ms-text-size-adjust:100%;-webkit-text-size-adjust:100%}.ExternalClass{width:100%}div[style*="margin: 16px 0"]{margin:0 !important}table,td{mso-table-lspace:0pt !important;mso-table-rspace:0pt !important}table{border-spacing:0 !important;border-collapse:collapse !important;table-layout:fixed !important;margin:0 auto !important}table table table{table-layout:auto}img{-ms-interpolation-mode:bicubic}.yshortcuts a{border-bottom:none !important}a[x-apple-data-detectors]{color:inherit !important}</style><style type="text/css">.button-td,.button-a{transition:all 100ms ease-in}.button-td:hover,.button-a:hover{background:#555 !important;border-color:#555 !important}@media screen and (max-width: 700px){.email-container{width:100% !important}.fluid,.fluid-centered{max-width:100% !important;height:auto !important;margin-left:auto !important;margin-right:auto !important}.fluid-centered{margin-left:auto !important;margin-right:auto !important}.stack-column,.stack-column-center{display:block !important;width:100% !important;max-width:100% !important;direction:ltr !important}.stack-column-center{text-align:center !important}.center-on-narrow{text-align:center !important;display:block !important;margin-left:auto !important;margin-right:auto !important;float:none !important}table.center-on-narrow{display:inline-block !important}}</style></head><body bgcolor="#fff" width="100%" style="margin: 0;" yahoo="yahoo"><table bgcolor="#fff" cellpadding="0" cellspacing="0" border="0" height="100%" width="100%" style="border-collapse:collapse;"><tr><td><center style="width: 100%;"><div style="display:none;font-size:1px;line-height:1px;max-height:0px;max-width:0px;opacity:0;overflow:hidden;mso-hide:all;font-family: sans-serif;">' + data.funcColum[0].columTitle + ': ' + data.funcColum[0].columText + '</div><table align="center" width="700" class="email-container"><tr bgcolor="#C9151E"><td style="padding: 0; text-align: center;"><a style="display: block; line-height: 0;" href="https://english.cw.com.tw/" target="_blank" rel="noopener noreferrer"><img src="https://topic.cw.com.tw/edm/cwEnglish/images/logo.jpg" style="display: inline-block; width: 140px;" alt="CommonWealth Magazine"></a></td></tr><tr><td style="padding: 1rem 0; text-align: center; border-bottom: 1px solid #c9c9c9;"><p style="margin: 0 0 5px; color:#333; font-family: Georgia, \'Libre Baskerville\', serif; font-size: 1rem; line-height: 1.5;"><span style="color: #999; font-size: 1rem;">2017-08-22 Tuesday</span><br/>' + data.funcColum[0].columTitle + '</p><p style="margin: 0; font-family: Georgia, \'Libre Baskerville\', serif;"><span style="color:#999; font-size: 0.8125rem;">Not Coming Through? Click <a id="edmLink" style="color:#D60C18; text-decoration: underline;" href="' + initial + "?file=" + key + '">here</a></span></p></td></tr><tr><td style="padding: 1rem 2.858%; border-bottom: 1px solid #c9c9c9;"><a href="' + data.funcColum[1].columLink + '" style="display: block; color: #666; font-family: Georgia, \'Libre Baskerville\', serif; font-size: 0.8125rem; line-height: 1.5; text-decoration: none;" target="_blank">' + data.funcColum[1].columTitle + '</a></td></tr><tr><td style="padding: 0.5rem 2.858%; border-bottom: 1px solid #c9c9c9;"><a href="' + initial + "?file=" + key + '&amp;playID=all" title="立刻聽天下" style="display: block; text-align: right; color: #666; font-family: Georgia, \'Libre Baskerville\', serif; font-size: 0.8125rem; line-height: 1.5; text-decoration: none;"><img style="display: inline-block; vertical-align: middle;" src="https://topic.cw.com.tw/edm/cwdaily/images/headphone@2x.png" width="20" alt="headphone icon"><span style="display: inline-block; font-family: Georgia, \'Libre Baskerville\', serif; vertical-align: middle;">&nbsp;&nbsp;Audio</span></a></td></tr><tr><td style="padding-top: 20px;"><iframe id="player" width="100%" height="280" scrolling="no" frameborder="no" src="' + playlist + "?file=" + key + '&Daily=' + data.funcColum[0].columChannel +'&playID=all' + '" style="display: none;"></iframe><div style="display: none; font-size: 0.75rem; text-align:left; margin: 5px 0 0 0;"><a href="https://vr2.cyberon.com.tw/cloud_tts/register/" style="text-decoration:none; color:#999">本網頁線上語音由賽微科技股份有限公司提供</a></div></td></tr></table><table cellspacing="0" cellpadding="0" border="0" align="center" bgcolor="#fff" width="700" class="email-container">';

        previewHead = '<table bgcolor="#fff" cellpadding="0" cellspacing="0" border="0" height="100%" width="100%" style="border-collapse:collapse;"><tr><td><center style="width: 100%;"><div style="display:none;font-size:1px;line-height:1px;max-height:0px;max-width:0px;opacity:0;overflow:hidden;mso-hide:all;font-family: sans-serif;">' + data.funcColum[0].columTitle + ': ' + data.funcColum[0].columText + '</div><table align="center" width="700" class="email-container"><tr bgcolor="#C9151E"><td style="padding: 0; text-align: center;"><a style="display: block; line-height: 0;" href="https://english.cw.com.tw/" target="_blank" rel="noopener noreferrer"><img src="https://topic.cw.com.tw/edm/cwEnglish/images/logo.jpg" style="display: inline-block; width: 140px;" alt="CommonWealth Magazine"></a></td></tr><tr><td style="padding: 1rem 0; text-align: center; border-bottom: 1px solid #c9c9c9;"><p style="margin: 0 0 5px; color:#333; font-family: Georgia, \'Libre Baskerville\', serif; font-size: 1rem; line-height: 1.5;"><span style="color: #999; font-size: 1rem;">2017-08-22 Tuesday</span><br/>' + data.funcColum[0].columTitle + '</p><p style="margin: 0; font-family: Georgia, \'Libre Baskerville\', serif;"><span style="color:#999; font-size: 0.8125rem;">Not Coming Through? Click <a id="edmLink" style="color:#D60C18; text-decoration: underline;" href="' + initial + "?file=" + key + '">here</a></span></p></td></tr><tr><td style="padding: 1rem 2.858%; border-bottom: 1px solid #c9c9c9;"><a href="' + data.funcColum[1].columLink + '" style="display: block; color: #666; font-family: Georgia, \'Libre Baskerville\', serif; font-size: 0.8125rem; line-height: 1.5; text-decoration: none;" target="_blank">' + data.funcColum[1].columTitle + '</a></td></tr><tr><td style="padding: 0.5rem 2.858%; border-bottom: 1px solid #c9c9c9;"><a href="' + initial + "?file=" + key + '&Daily=' + data.funcColum[0].columChannel +'&amp;playID=all" title="立刻聽天下" style="display: block; text-align: right; color: #666; font-family: Georgia, \'Libre Baskerville\', serif; font-size: 0.8125rem; line-height: 1.5; text-decoration: none;"><img style="display: inline-block; vertical-align: middle;" src="https://topic.cw.com.tw/edm/cwdaily/images/headphone@2x.png" width="20" alt="headphone icon"><span style="display: inline-block; font-family: Georgia, \'Libre Baskerville\', serif; vertical-align: middle;">&nbsp;&nbsp;Audio</span></a></td></tr><tr><td style="padding-top: 20px;"><iframe id="player" width="100%" height="280" scrolling="no" frameborder="no" src="' + playlist + "?file=" + key + '&playID=all' + '" style="display: none;"></iframe><div style="display: none; font-size: 0.75rem; text-align:left; margin: 5px 0 0 0;"><a href="https://vr2.cyberon.com.tw/cloud_tts/register/" style="text-decoration:none; color:#999">本網頁線上語音由賽微科技股份有限公司提供</a></div></td></tr></table><table cellspacing="0" cellpadding="0" border="0" align="center" bgcolor="#fff" width="700" class="email-container">';

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
                    case 'article':
                        console.log('Double-column, ' + multiColumCont);
                        codeBody += '<tr><td dir="ltr" align="center" valign="top" width="100%" style="padding: 20px 1.43% 10px; border-bottom: 1px solid #c9c9c9;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td style="padding: 0 10px 10px;"><span style="display: block; color: #D60C18; font-family: Georgia, \'Libre Baskerville\', serif; font-size: 1rem; font-weight: bold;">» ' + f.columChannel + '</span></td></tr></table><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td width="44%" valign="top" class="stack-column-center"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td dir="ltr" valign="top" style="padding: 0 10px 10px; line-height: 0"><a href="' + f.columLink + '" target="_blank" style="display: block;"><img src="' + f.columImage + '" width="100%" alt="alt_text" border="0" class="center-on-narrow"></a></td></tr></table></td><td width="56%" valign="top" class="stack-column-center"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td dir="ltr" valign="top" style="padding: 0 10px 10px;" class="center-on-narrow"><a href="' + f.columLink + '" style="display: block; padding: 0 0 10px; text-align: left; text-decoration: none;" target="_blank"><h3 style="margin: 0; color: #1f4f82; font-family: Georgia, \'Libre Baskerville\', serif; font-size: 1.25rem; font-weight: bold; mso-height-rule: exactly; line-height: 1.5;">' + f.columTitle + '</h3></a><p style="margin: 0; padding: 0; color: #333; font-family: Georgia, \'Libre Baskerville\', serif; font-size: 0.875rem; mso-height-rule: exactly; text-align: left; line-height: 1.5;">' + f.columText + ' <a href="' + f.columLink + '" style="text-decoration: none;" target="_blank"><img src="https://topic.cw.com.tw/edm/cwEnglish/images/more.png" width="40" style="display: inline-block; vertical-align: middle;" alt="read more"></a></p></td></tr></table></td></tr></table></td></tr>';
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
        codeFoot = '</table><table cellspacing="0" cellpadding="0" border="0" align="center" bgcolor="#545454" width="700" class="email-container"><tr><td style="padding: 1rem 2.858%; text-align: center;"><table cellspacing="0" cellpadding="0" border="0" width="100%" align="center" class="email-container"><tr><td align="center" style="padding-top: 0; font-size: 0; font-family: Georgia, \'Libre Baskerville\', serif"><a href="https://www.cw.com.tw/article/articleLogin.action?id=5096027&utm_source=email_edm&utm_medium=email" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 0.75rem; color: #fff; text-decoration: none;">About Us</a><a href="https://account.cwg.tw/register/cw?utm_source=email_edm&utm_medium=email" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 0.75rem; color: #fff; text-decoration: none; border-left: 1px solid #fff;">Member Service</a><a href="https://www.cw.com.tw/saleskit/" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 0.75rem; color: #fff; text-decoration: none; border-left: 1px solid #fff;">Advertising With Us</a><a href="https://www.cw.com.tw/recruit/join1.php" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 0.75rem; color: #fff; text-decoration: none; border-left: 1px solid #fff;">Work With Us</a><a href="https://www.cw.com.tw/member/newsletters?utm_source=email_edm&utm_medium=email" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 0.75rem; color: #fff; text-decoration: none; border-left: 1px solid #fff;">Unsubscribe</a></td></tr><tr><td align="center" style="padding: 0.5rem 0; color: #fff; font-size: 0.75rem; font-family: Georgia, \'Libre Baskerville\', serif">Copyright © ' + currentYear + ' CommonWealth Magazine All rights reserved.</td></tr><tr><td style="font-size: 0;"><a href="https://www.facebook.com/CommonWealthMagazine/?fref=ts" target="_blank" rel="noopener noreferrer" style="display: inline-block; margin: 0; padding: 0;" title="天下雜誌 Facebook 粉絲專頁"><img src="https://topic.cw.com.tw/edm/cwEnglish/images/fb.jpg" width="80" alt="facebook"></a><a href="https://line.me/ti/p/@xat.0000135614.9o8" target="_blank" rel="noopener noreferrer" style="display: inline-block; margin: 0; padding: 0;" title="天下雜誌 Line"><img src="https://topic.cw.com.tw/edm/cwEnglish/images/line.jpg" width="80" alt="line"></a></td></tr></table></td></tr></table><custom name="opencounter" type="tracking"/><div style="display: none; font-size: 1px; line-height: 1px; max-height: 0px; max-width: 0px; opacity: 0; overflow: hidden; mso-hide: all; font-family: sans-serif;">This email was sent by:<b>%%Member_Busname%%</b><br>%%Member_Addr%% %%Member_City%%, %%Member_State%%, %%Member_PostalCode%%, %%Member_Country%%</div></center></td></tr></table></body></html>';

        previewFoot = '</table><table cellspacing="0" cellpadding="0" border="0" align="center" bgcolor="#545454" width="700" class="email-container"><tr><td style="padding: 1rem 2.858%; text-align: center;"><table cellspacing="0" cellpadding="0" border="0" width="100%" align="center" class="email-container"><tr><td align="center" style="padding-top: 0; font-size: 0; font-family: Georgia, \'Libre Baskerville\', serif"><a href="https://www.cw.com.tw/article/articleLogin.action?id=5096027&utm_source=email_edm&utm_medium=email" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 0.75rem; color: #fff; text-decoration: none;">About Us</a><a href="https://account.cwg.tw/register/cw?utm_source=email_edm&utm_medium=email" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 0.75rem; color: #fff; text-decoration: none; border-left: 1px solid #fff;">Member Service</a><a href="https://www.cw.com.tw/saleskit/" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 0.75rem; color: #fff; text-decoration: none; border-left: 1px solid #fff;">Advertising With Us</a><a href="https://www.cw.com.tw/recruit/join1.php" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 0.75rem; color: #fff; text-decoration: none; border-left: 1px solid #fff;">Work With Us</a><a href="https://www.cw.com.tw/member/newsletters?utm_source=email_edm&utm_medium=email" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; font-size: 0.75rem; color: #fff; text-decoration: none; border-left: 1px solid #fff;">Unsubscribe</a></td></tr><tr><td align="center" style="padding: 0.5rem 0; color: #fff; font-size: 0.75rem; font-family: Georgia, \'Libre Baskerville\', serif">Copyright © ' + currentYear + ' CommonWealth Magazine All rights reserved.</td></tr><tr><td style="font-size: 0;"><a href="https://www.facebook.com/CommonWealthMagazine/?fref=ts" target="_blank" rel="noopener noreferrer" style="display: inline-block; margin: 0; padding: 0;" title="天下雜誌 Facebook 粉絲專頁"><img src="https://topic.cw.com.tw/edm/cwEnglish/images/fb.jpg" width="80" alt="facebook"></a><a href="https://line.me/ti/p/@xat.0000135614.9o8" target="_blank" rel="noopener noreferrer" style="display: inline-block; margin: 0; padding: 0;" title="天下雜誌 Line"><img src="https://topic.cw.com.tw/edm/cwEnglish/images/line.jpg" width="80" alt="line"></a></td></tr></table></td></tr></table><custom name="opencounter" type="tracking"/><div style="display: none; font-size: 1px; line-height: 1px; max-height: 0px; max-width: 0px; opacity: 0; overflow: hidden; mso-hide: all; font-family: sans-serif;">This email was sent by:<b>%%Member_Busname%%</b><br>%%Member_Addr%% %%Member_City%%, %%Member_State%%, %%Member_PostalCode%%, %%Member_Country%%</div></center></td></tr></table>';

        // dataLayer.push({
        //     "event": "GA-event",
        //     "eventInfo": "編碼完成",
        //     "fileID": key
        // });
        //將原始碼塞入「發信用原始碼」的框框
        $("#sourceCode textarea").val(codeHead+codeBody+codeFoot);
        //將原始碼丟入「預覽」框，呈現組裝後的結果
        $("#previewHtml").html(previewHead+codeBody+previewFoot);
        $("#initial").html(previewHead+codeBody+previewFoot);
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
    $.getJSON("https://spreadsheets.google.com/feeds/list/" + key + "/1/public/values?alt=json&sq=display=y", function(l) {
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
        document.title = data.funcColum[0].columTitle + " - CommonWealth Magazine Weekly Newsletter";
        // ga('send', 'event', 'edmcode', 'finish', 'make-edm success');
        if ( window.location.href.indexOf('playID') > 0 ) {
            $('#player').show();
            $('#player').siblings().show();
        }
    })
}
