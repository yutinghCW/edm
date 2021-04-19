var width = $(window).width(),
	initial = "https://yutinghcw.github.io/edm/initial/opinion-v2.html",
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
	$.getJSON(gspreadsheets, function(result){
		//將載入列表轉換為清單的格式
		$.each(result.feed.entry, function(i, f) {
			data.funcColum[i] = {
				columType: f.gsx$欄型.$t,
				columLink: f.gsx$連結.$t,
				columChannel: f.gsx$頻道.$t,
				columChannelLink: f.gsx$頻道連結.$t,
				columAuthor: f.gsx$作者.$t,
				columButton: f.gsx$按鈕文字.$t,
				columImage: f.gsx$圖片.$t,
				columTitle: f.gsx$標題.$t,
				columText: f.gsx$內文.$t,
				columNote: f.gsx$標誌.$t
			};
		});

		// 組裝原始碼到body前
		codeHead = '<!doctype html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><meta http-equiv="X-UA-Compatible" content="IE=edge"><title>EmailTemplate-Responsive</title><link rel="preconnect" href="https://fonts.gstatic.com"><link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500&family=Noto+Serif+TC:wght@500;600;700&family=Roboto+Slab:wght@400;500&family=Roboto:wght@300;400;500&display=swap" rel="stylesheet"><style type="text/css">html,body{margin:0 !important;padding:0 !important;height:100% !important;width:100% !important}*{-ms-text-size-adjust:100%;-webkit-text-size-adjust:100%}.ExternalClass{width:100%}div[style*="margin: 16px 0"]{margin:0 !important}table,td{mso-table-lspace:0pt !important;mso-table-rspace:0pt !important}table{border-spacing:0 !important;border-collapse:collapse !important;table-layout:fixed !important;margin:0 auto !important}table table table{table-layout:auto}img{-ms-interpolation-mode:bicubic}.yshortcuts a{border-bottom:none !important}a[x-apple-data-detectors]{color:inherit !important}</style><style type="text/css">.button-td,.button-a{transition:all 100ms ease-in}.button-td:hover,.button-a:hover{background:#20204c !important;border-color:#20204c !important}@media screen and (max-width: 600px){.email-container{width:100% !important}.fluid,.fluid-centered{max-width:100% !important;height:auto !important;margin-left:auto !important;margin-right:auto !important}.fluid-centered{margin-left:auto !important;margin-right:auto !important}.stack-column,.stack-column-center{display:block !important;width:100% !important;max-width:100% !important;direction:ltr !important}.stack-column-center{text-align:center !important}.center-on-narrow{text-align:center !important;display:block !important;margin-left:auto !important;margin-right:auto !important;float:none !important}table.center-on-narrow{display:inline-block !important}}</style></head><body bgcolor="#F2F2F2" width="100%" style="margin: 0;" yahoo="yahoo"><table bgcolor="#F2F2F2" cellpadding="0" cellspacing="0" border="0" height="100%" width="100%" style="border-collapse:collapse;"><tr><td style="padding: 20px;"><center style="width: 100%;">';

		// 組裝原始碼第一段，版頭及表頭
		previewHead = '<div style="display:none; font-size:1px; line-height:1px; max-height:0px; max-width:0px; opacity:0; overflow:hidden; mso-hide:all; font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif;">我是一個大家看不到的前言</div><table cellspacing="0" cellpadding="0" border="0" align="center" width="600" class="email-container unsubscribe"><tr><td style="padding: 0 0 10px; text-align: right;"><a href="' + initial + "?file=" + key + '" target="_blank" rel="noopener noreferrer" style="display: block; color: #777; font-size: 0.875rem">無法看到完整內容，請點此處</a></td></tr></table><table cellspacing="0" cellpadding="0" border="0" align="center" bgcolor="#ffffff" width="600" class="email-container"><tr><td style="padding: 20px 6.6675%; line-height: 0; text-align: center;"><a href="http://opinion.cw.com.tw/" target="_blank" rel="noopener noreferrer"><img src="https://topic.cw.com.tw/salesforce/assets/images/opinion-logo-vertical@3x.png" height="55" alt="獨立評論@天下 Logo" border="0" style="display: inline-block;"></a></td></tr><tr><td style="text-align: center"><h2 style="margin: 0 0 5px; padding: 0 6.6675%; color: #313160; font-family: \'Noto Serif TC\', \'思源宋體 TC\', \'思源宋體 TW\', \'思源宋體\', serif; font-size: 1.75rem;">' + data.funcColum[0].columTitle + '</h2><time style="margin-bottom: 10px; padding: 0 6.6675%; color: #777; font-size: 0.875rem;">' + data.funcColum[0].columLink + '</time><div style="margin: 10px 0 20px; padding: 0 6.6675%; color: #000; font-size: 1rem; font-weight: 400; line-height: 1.7; text-align: left; white-space: pre-line;">' + data.funcColum[0].columText + '</div>';

		if ( data.funcColum[0].columButton !== "" ) {
			previewHead += '<table cellspacing="0" cellpadding="0" border="0" align="center"><tr><td style="text-align: center;"><a href="' + data.funcColum[0].columChannelLink + '" style="display: block; padding: 0 1em; color: #ffffff; font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif; font-size: 0.875rem; text-decoration: none; background: #313160; border: 0.5em solid #313160; border-radius: 5px;" class="button-a"><!--[if mso]>&nbsp;&nbsp;&nbsp;&nbsp;<![endif]-->' + data.funcColum[0].columButton + '<!--[if mso]>&nbsp;&nbsp;&nbsp;&nbsp;<![endif]--></a></td></tr></table>';
		}
		if ( data.funcColum[0].columImage == "" ) {
			previewHead += '<div style="padding: 20px 0 0;"></div>';
		} else {
			previewHead += '<img src="' + data.funcColum[0].columImage + '" style="display: block; margin-top: 20px;" width="100%" border="0" alt="">';
		}
		previewHead += '</td></tr></table>';

		// 組裝原始碼第二段，填入spreadsheet的內容部份
		var multiColumCont = 0;
		$(function() {
			$.each(data.funcColum, function(i, f) {
				if (i == 0) {
					return true;
				}
				switch (f.columType) {
					case 'Block-Start':
						codeBody += '<table cellspacing="0" cellpadding="0" border="0" align="center" bgcolor="#ffffff" width="600" class="email-container">';
						// console.log('Block-Start');
						break;
					case 'Block-End':
						codeBody += '</table><div style="padding: 20px 0;"></div>';
						// console.log('Block-End');
						break;
					case 'Block-Title':
						if ( ((i-1) == 1) && (data.funcColum[0].columText == "") && (data.funcColum[0].columButton == "") && (data.funcColum[0].columImage == "") ) {
							codeBody += '<tr><td style="padding: 0 6.667%;"><table cellspacing="0" cellpadding="0" border="0" align="left"><tr><td style="padding: 0 0 0 20px; color: #313160; font-size: 22px; font-family: \'Noto Serif TC\', \'思源宋體 TC\', \'思源宋體 TW\', \'思源宋體\', serif; font-weight: 600; border-left: 4px solid #313160;">' + f.columTitle + '</td></tr></table></td></tr>';
						} else {
							codeBody += '<tr><td style="padding: 0 6.667%;"><table cellspacing="0" cellpadding="0" border="0" align="left"><tr><td style="padding: 40px 0 0 20px; color: #313160; font-size: 22px; font-family: \'Noto Serif TC\', \'思源宋體 TC\', \'思源宋體 TW\', \'思源宋體\', serif; font-weight: 600; border-left: 4px solid #313160;">' + f.columTitle + '</td></tr></table></td></tr>';
						}
						// console.log('Block-Title');
						break;
					case 'Single-Article':
						codeBody += '<tr><td style="padding: 20px 6.667% 0;"><a href="' + f.columLink + '" target="_blank" rel="noopener noreferrer" style="display: block;"><img src="' + f.columImage + '" width="600" alt="alt_text" border="0" align="center" style="width: 100%; max-width: 600px; height: auto;"></a></td></tr><tr><td dir="ltr" align="center" valign="top" width="100%" style="padding: 10px 6.667%;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td style="padding: 0 10% 10px;"><a href="' + f.columLink + '" target="_blank" rel="noopener noreferrer" style="display: block; color: #000; font-size: 1.125rem; line-height: 1.5; font-weight: 500; text-decoration: none;">' + f.columTitle + '</a>';
						if ( f.columText !== "" ) {
							var TextEllipsis_55_1 = "";
							if ( f.columText.length > 55 ) {
								TextEllipsis_55_1 = f.columText.substring(0, 54) + "⋯⋯";
							} else {
								TextEllipsis_55_1 = f.columText;
							}
							codeBody += '<div style="padding: 5px 0; color: #777; font-size: 0.875rem; font-weight: 400; line-height: 1.25; white-space: pre-line;">' + TextEllipsis_55_1 + '</div>';
						}
						if ( (f.columChannel !== "") || (f.columAuthor !== "") ) {
							codeBody += '<div style="padding: 5px 0; color: #777; font-size: 0.875rem; line-height: 1;">';
							if ( f.columChannel !== "" ) {
								codeBody += '<a href="' + f.columChannelLink + '" target="_blank" rel="noopener noreferrer" style="margin-right: 10px; color: #313160; font-weight: 500; text-decoration: none;">' + f.columChannel + '</a>';
							}
							if ( f.columAuthor !== "" ) {
								codeBody += '<span style="font-weight: 400;">作者 ' + f.columAuthor + '</span>';
							}
							codeBody += '</div>';
						}
						if ( f.columButton !== "" ) {
							codeBody += '<table cellspacing="0" cellpadding="0" border="0" align="center"><tr><td><a href="' + f.columLink + '" style="display: block; padding: 0 1em; margin: 20px 0 0; color: #ffffff;font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif; font-size: 0.875rem; text-decoration: none; background: #313160; border: 0.5em solid #313160; border-radius: 5px;" class="button-a"><!--[if mso]>&nbsp;&nbsp;&nbsp;&nbsp;<![endif]-->' + f.columButton + '<!--[if mso]>&nbsp;&nbsp;&nbsp;&nbsp;<![endif]--></a></td></tr></table>';
						}
						codeBody += '</td></tr></table></td></tr>';
						if ( f.columNote == "" ) {
							codeBody += '<tr><td style="padding: 0 6.667%;"><table cellspacing="0" cellpadding="0" border="0" align="center" bgcolor="#dadada" width="100%" height="1"><tr><td></td></tr></table></td></tr>';
						}
						// console.log('Single-Article');
						break;
					case 'Text-Button':
						codeBody += '<tr><td dir="ltr" align="center" valign="top" width="100%" style="padding: 10px 6.667%;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td class="stack-column" style="padding: 10px 0;"><a href="' + f.columLink + '" target="_blank" rel="noopener noreferrer" style="display: block; color: #000; font-size: 1.125rem; line-height: 1.5; font-weight: 500; text-decoration: none;">' + f.columTitle + '</a>';
						if ( f.columText !== "" ) {
							var TextEllipsis_55_2 = "";
							if ( f.columText.length > 55 ) {
								TextEllipsis_55_2 = f.columText.substring(0, 54) + "⋯⋯";
							} else {
								TextEllipsis_55_2 = f.columText;
							}
							codeBody += '<div style="padding: 5px 0; color: #777; font-size: 0.875rem; font-weight: 400; line-height: 1.25; white-space: pre-line;">' + TextEllipsis_55_2 + '</div>';
						}
						codeBody += '<div style="padding: 5px 0; color: #777; font-size: 0.875rem; line-height: 1;"><a href="' + f.columChannelLink + '" target="_blank" rel="noopener noreferrer" style="margin-right: 10px; color: #313160; font-weight: 500; text-decoration: none;">' + f.columChannel + '</a><span style="font-weight: 400;">作者 ' + f.columAuthor + '</span></div></td><td class="stack-column" width="20"></td><td class="stack-column" style="width: 6.125rem;"><table cellspacing="0" cellpadding="0" border="0" align="center"><tr><td><a href="' + f.columLink + '" style="display: block; padding: 0 1em; margin: 10px 0; color: #ffffff;font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif; font-size: 0.875rem; text-decoration: none; background: #313160; border: 0.5em solid #313160; border-radius: 5px;" class="button-a"><!--[if mso]>&nbsp;&nbsp;&nbsp;&nbsp;<![endif]-->' + f.columButton + '<!--[if mso]>&nbsp;&nbsp;&nbsp;&nbsp;<![endif]--></a></td></tr></table></td></tr></table></td></tr>';
						if ( f.columNote == "" ) {
							codeBody += '<tr><td style="padding: 0 6.667%;"><table cellspacing="0" cellpadding="0" border="0" align="center" bgcolor="#dadada" width="100%" height="1"><tr><td></td></tr></table></td></tr>';
						}
						// console.log('Text-Button');
						break;
					case 'Ract-Img-Text':
						const ResizeImg_240by150 = f.columImage.split('w=')[0] + 'w=240&h=150&fit=cover';
						codeBody += '<tr><td dir="ltr" align="center" valign="top" width="100%" style="padding: 10px 6.667%;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td class="stack-column" width="240" style="padding: 10px 0;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td dir="ltr" valign="top"><a href="' + f.columLink + '" target="_blank" rel="noopener noreferrer" style="display: block;"><img src="' + ResizeImg_240by150 + '" width="100%" alt="" border="0" style="display: block;"></a></td></tr></table></td><td class="stack-column" width="20"></td><td class="stack-column" style="padding: 10px 0;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td dir="ltr" valign="top" style="font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif; font-size: 15px; mso-height-rule: exactly; text-align: left;"><a href="' + f.columLink + '" target="_blank" rel="noopener noreferrer" style="display: block; color: #000; font-size: 1.125rem; line-height: 1.5; font-weight: 500; text-decoration: none;">' + f.columTitle + '</a>';
						if ( f.columText !== "" ) {
							var TextEllipsis_55_3 = "";
							if ( f.columText.length > 55 ) {
								TextEllipsis_55_3 = f.columText.substring(0, 54) + "⋯⋯";
							} else {
								TextEllipsis_55_3 = f.columText;
							}
							codeBody += '<div style="padding: 5px 0; color: #777; font-size: 0.875rem; font-weight: 400; line-height: 1.25; white-space: pre-line;">' + TextEllipsis_55_3 + '</div>';
						}
						if ( (f.columChannel !== "") || (f.columAuthor !== "") ) {
							codeBody += '<div style="padding: 5px 0; color: #777; font-size: 0.875rem; line-height: 1;">';
							if ( f.columChannel !== "" ) {
								codeBody += '<a href="' + f.columChannelLink + '" target="_blank" rel="noopener noreferrer" style="margin-right: 10px; color: #313160; font-weight: 500; text-decoration: none;">' + f.columChannel + '</a>';
							}
							if ( f.columAuthor !== "" ) {
								codeBody += '<span style="font-weight: 400;">作者 ' + f.columAuthor + '</span>';
							}
							codeBody += '</div>';
						}
						if ( f.columButton !== "" ) {
							codeBody += '<table cellspacing="0" cellpadding="0" border="0" width="100%"><tr><td class="stack-column-center" style="text-align: right;"><a href="' + f.columLink + '" target="_blank" rel="noopener noreferrer" style="display: inline-block; padding: 0 1em;margin: 15px 0 0;color: #ffffff;font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif;font-size: 0.875rem;text-decoration: none;background: #313160;border: 0.5em solid #313160;border-radius: 5px;" class="button-a"><!--[if mso]>&nbsp;&nbsp;&nbsp;&nbsp;<![endif]-->' + f.columButton + '<!--[if mso]>&nbsp;&nbsp;&nbsp;&nbsp;<![endif]--></a></td></tr></table>';
						}
						codeBody += '</td></tr></table></td></tr></table></td></tr>';
						if ( f.columNote == "" ) {
							codeBody += '<tr><td style="padding: 0 6.667%;"><table cellspacing="0" cellpadding="0" border="0" align="center" bgcolor="#dadada" width="100%" height="1"><tr><td></td></tr></table></td></tr>';
						}
						// console.log('Ract-Img-Text');
						break;
					case 'Medium-Img-Text':
						codeBody += '<tr><td dir="ltr" align="center" valign="top" width="100%" style="padding: 10px 6.667%;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td class="stack-column" width="180" style="padding: 10px 0;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td dir="ltr" valign="top"><a href="' + f.columLink + '" target="_blank" rel="noopener noreferrer" style="display: block;"><img src="' + f.columImage + '" width="100%" alt="" border="0" style="display: block;"></a></td></tr></table></td><td class="stack-column" width="20"></td><td class="stack-column" style="padding: 10px 0;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td dir="ltr" valign="top" style="font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif; font-size: 15px; mso-height-rule: exactly; text-align: left;"><a href="' + f.columLink + '" target="_blank" rel="noopener noreferrer" style="display: block; color: #000; font-size: 1.125rem; line-height: 1.5; font-weight: 500; text-decoration: none;">' + f.columTitle + '</a>';
						if ( f.columText !== "" ) {
							var TextEllipsis_55_3 = "";
							if ( f.columText.length > 55 ) {
								TextEllipsis_55_3 = f.columText.substring(0, 54) + "⋯⋯";
							} else {
								TextEllipsis_55_3 = f.columText;
							}
							codeBody += '<div style="padding: 5px 0; color: #777; font-size: 0.875rem; font-weight: 400; line-height: 1.25; white-space: pre-line;">' + TextEllipsis_55_3 + '</div>';
						}
						if ( (f.columChannel !== "") || (f.columAuthor !== "") ) {
							codeBody += '<div style="padding: 5px 0; color: #777; font-size: 0.875rem; line-height: 1;">';
							if ( f.columChannel !== "" ) {
								codeBody += '<a href="' + f.columChannelLink + '" target="_blank" rel="noopener noreferrer" style="margin-right: 10px; color: #313160; font-weight: 500; text-decoration: none;">' + f.columChannel + '</a>';
							}
							if ( f.columAuthor !== "" ) {
								codeBody += '<span style="font-weight: 400;">作者 ' + f.columAuthor + '</span>';
							}
							codeBody += '</div>';
						}
						if ( f.columButton !== "" ) {
							codeBody += '<table cellspacing="0" cellpadding="0" border="0" width="100%"><tr><td class="stack-column-center" style="text-align: right;"><a href="' + f.columLink + '" target="_blank" rel="noopener noreferrer" style="display: inline-block; padding: 0 1em;margin: 15px 0 0;color: #ffffff;font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif;font-size: 0.875rem;text-decoration: none;background: #313160;border: 0.5em solid #313160;border-radius: 5px;" class="button-a"><!--[if mso]>&nbsp;&nbsp;&nbsp;&nbsp;<![endif]-->' + f.columButton + '<!--[if mso]>&nbsp;&nbsp;&nbsp;&nbsp;<![endif]--></a></td></tr></table>';
						}
						codeBody += '</td></tr></table></td></tr></table></td></tr>';
						if ( f.columNote == "" ) {
							codeBody += '<tr><td style="padding: 0 6.667%;"><table cellspacing="0" cellpadding="0" border="0" align="center" bgcolor="#dadada" width="100%" height="1"><tr><td></td></tr></table></td></tr>';
						}
						// console.log('Medium-Img-Text');
						break;
					case 'Circle-Img-Left':
						const ResizeImg_150by150_1 = f.columImage.split('w=')[0] + 'w=150&h=150&fit=cover';
						codeBody += '<tr><td dir="ltr" align="center" valign="top" width="100%" style="padding: 10px 6.667%;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td class="stack-column" width="135" style="padding: 10px 0;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td dir="ltr" valign="top"><a href="' + f.columLink + '" target="_blank" rel="noopener noreferrer" style="display: block;"><img src="' + ResizeImg_150by150_1 + '" width="100%" alt="" border="0" style="display: block; max-width: 150px; margin: auto; border-radius: 50%;"></a></td></tr></table></td><td class="stack-column" width="20"></td><td class="stack-column" style="padding: 10px 0;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td dir="ltr" valign="top" style="font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif; font-size: 15px; mso-height-rule: exactly; text-align: left;"><a href="' + f.columLink + '" target="_blank" rel="noopener noreferrer" style="color: #000; font-size: 1.125rem; line-height: 1.5; font-weight: 500; text-decoration: none;">' + f.columTitle + '</a>';
						if ( f.columText !== "" ) {
							var TextEllipsis_55_4 = "";
							if ( f.columText.length > 55 ) {
								TextEllipsis_55_4 = f.columText.substring(0, 54) + "⋯⋯";
							} else {
								TextEllipsis_55_4 = f.columText;
							}
							codeBody += '<div style="padding: 5px 0; color: #777; font-size: 0.875rem; font-weight: 400; line-height: 1.25; white-space: pre-line;">' + TextEllipsis_55_4 + '</div>';
						}
						if ( (f.columChannel !== "") || (f.columAuthor !== "") ) {
							codeBody += '<div style="padding: 5px 0; color: #777; font-size: 0.875rem; line-height: 1;">';
							if ( f.columChannel !== "" ) {
								codeBody += '<a href="' + f.columChannelLink + '" target="_blank" rel="noopener noreferrer" style="margin-right: 10px; color: #313160; font-weight: 500; text-decoration: none;">' + f.columChannel + '</a>';
							}
							if ( f.columAuthor !== "" ) {
								codeBody += '<span style="font-weight: 400;">作者 ' + f.columAuthor + '</span>';
							}
							codeBody += '</div>';
						}
						if ( f.columButton !== "" ) {
							codeBody += '<table cellspacing="0" cellpadding="0" border="0" width="100%"><tr><td class="stack-column-center"><a href="' + f.columLink + '" target="_blank" rel="noopener noreferrer" style="display: inline-block; padding: 0 1em;margin: 15px 0 0;color: #ffffff;font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif;font-size: 0.875rem;text-decoration: none;background: #313160;border: 0.5em solid #313160;border-radius: 5px;" class="button-a"><!--[if mso]>&nbsp;&nbsp;&nbsp;&nbsp;<![endif]-->' + f.columButton + '<!--[if mso]>&nbsp;&nbsp;&nbsp;&nbsp;<![endif]--></a></td></tr></table>';
						}
						codeBody += '</td></tr></table></td></tr></table></td></tr>';
						if ( f.columNote == "" ) {
							codeBody += '<tr><td style="padding: 0 6.667%;"><table cellspacing="0" cellpadding="0" border="0" align="center" bgcolor="#dadada" width="100%" height="1"><tr><td></td></tr></table></td></tr>';
						}
						// console.log('Circle-Img-Left');
						break;
					case 'Circle-Img-Right':
						const ResizeImg_150by150_2 = f.columImage.split('w=')[0] + 'w=150&h=150&fit=cover';
						codeBody += '<tr><td dir="rtl" align="center" valign="top" width="100%" style="padding: 10px 6.667%;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td class="stack-column" width="135" style="padding: 10px 0;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td dir="ltr" valign="top"><a href="' + f.columLink + '" target="_blank" rel="noopener noreferrer" style="display: block;"><img src="' + ResizeImg_150by150_2 + '" width="100%" alt="" border="0" style="display: block; max-width: 150px; margin: auto; border-radius: 50%;"></a></td></tr></table></td><td class="stack-column" width="20"></td><td class="stack-column" style="padding: 10px 0;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td dir="ltr" valign="top" style="font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif; font-size: 15px; mso-height-rule: exactly; text-align: left;"><a href="' + f.columLink + '" target="_blank" rel="noopener noreferrer" style="color: #000; font-size: 1.125rem; line-height: 1.5; font-weight: 500; text-decoration: none;">' + f.columTitle + '</a>';
						if ( f.columText !== "" ) {
							var TextEllipsis_55_5 = "";
							if ( f.columText.length > 55 ) {
								TextEllipsis_55_5 = f.columText.substring(0, 54) + "⋯⋯";
							} else {
								TextEllipsis_55_5 = f.columText;
							}
							codeBody += '<div style="padding: 5px 0; color: #777; font-size: 0.875rem; font-weight: 400; line-height: 1.25; white-space: pre-line;">' + TextEllipsis_55_5 + '</div>';
						}
						if ( (f.columChannel !== "") || (f.columAuthor !== "") ) {
							codeBody += '<div style="padding: 5px 0; color: #777; font-size: 0.875rem; line-height: 1;">';
							if ( f.columChannel !== "" ) {
								codeBody += '<a href="' + f.columChannelLink + '" target="_blank" rel="noopener noreferrer" style="margin-right: 10px; color: #313160; font-weight: 500; text-decoration: none;">' + f.columChannel + '</a>';
							}
							if ( f.columAuthor !== "" ) {
								codeBody += '<span style="font-weight: 400;">作者 ' + f.columAuthor + '</span>';
							}
							codeBody += '</div>';
						}
						if ( f.columButton !== "" ) {
							codeBody += '<table cellspacing="0" cellpadding="0" border="0" width="100%"><tr><td class="stack-column-center"><a href="' + f.columLink + '" target="_blank" rel="noopener noreferrer" style="display: inline-block; padding: 0 1em;margin: 15px 0 0;color: #ffffff;font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif;font-size: 0.875rem;text-decoration: none;background: #313160;border: 0.5em solid #313160;border-radius: 5px;" class="button-a"><!--[if mso]>&nbsp;&nbsp;&nbsp;&nbsp;<![endif]-->' + f.columButton + '<!--[if mso]>&nbsp;&nbsp;&nbsp;&nbsp;<![endif]--></a></td></tr></table>';
						}
						codeBody += '</td></tr></table></td></tr></table></td></tr>';
						if ( f.columNote == "" ) {
							codeBody += '<tr><td style="padding: 0 6.667%;"><table cellspacing="0" cellpadding="0" border="0" align="center" bgcolor="#dadada" width="100%" height="1"><tr><td></td></tr></table></td></tr>';
						}
						break;
					case 'Rank':
						const ResizeImg_350by250 = f.columImage.split('w=')[0] + 'w=350&h=250&fit=cover';
						codeBody += '<tr><td><div style="padding-top: 20px;"></div></td></tr><tr><td>';
						if ( f.columChannel !== "" ) {
							codeBody += '<table cellspacing="0" cellpadding="0" border="0" align="center" width="240"><tr><td valign="bottom" style="width: 27.5%; color: #313160; font-family: \'Times New Roman\', Times, serif; font-style: italic; font-size: 7.5rem; line-height: 1; white-space: nowrap;">' + f.columChannel + '</td><td valign="bottom" style="width: 72.5%; padding: 0 0 1.25rem; color: #313160;"><div style="font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif; font-size: 1rem;">' + f.columChannelLink + '</div><div style="padding: 3px 0 0; font-family: \'Noto Serif TC\', \'思源宋體 TC\', \'思源宋體 TW\', \'思源宋體\', serif; font-weight: bold; font-size: 1.5rem; line-height: 1.125;">' + f.columTitle + '</div></td></tr></table>';
						} else {
							codeBody += '<table cellspacing="0" cellpadding="0" border="0" align="center" style="max-width: 240px;"><tr><td valign="bottom" style="padding: 0 0 0 20px; color: #313160; border-left: 4px solid #313160;"><div style="font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif; font-size: 1rem;">' + f.columChannelLink + '</div><div style="padding: 3px 0 0; font-family: \'Noto Serif TC\', \'思源宋體 TC\', \'思源宋體 TW\', \'思源宋體\', serif; font-weight: bold; font-size: 1.5rem; line-height: 1.125;">' + f.columTitle + '</div></td></tr><tr><td style="padding: 20px 0 0;"></td></tr></table>';
						}
						codeBody += '</td></tr><tr><td style="padding: 0 6.6667%;"><table cellspacing="0" cellpadding="0" border="0" align="center"><tr><td><a href="#!" target="_blank" rel="noopener noreferrer" style="display: block;"><img src="' + ResizeImg_350by250 + '" width="350" alt="alt_text" border="0" style="display: block; width: 100%; max-width: 350px; height: auto;"></a></td></tr><tr><td style="padding: 5px 0 20px; color: #777; font-size: 0.875rem; font-weight: 400;">' + f.columAuthor + '</td></tr></table></td></tr>';
						if ( f.columText !== "" ) {
							codeBody += '<tr><td style="padding: 0 6.6667%; color: #000; font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif; font-size: 1rem; font-weight: 400; mso-height-rule: exactly; line-height: 1.5; white-space: pre-line;">' + f.columText + '</td></tr>';
						}
						if ( f.columButton !== "" ) {
							codeBody += '<tr><td style="padding: 20px 0;"><table cellspacing="0" cellpadding="0" border="0" align="center"><tr><td style="text-align: center;"><a href="#!" style="display: block; padding: 0 1em; color: #ffffff; font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif; font-size: 0.875rem; text-decoration: none; background: #313160; border: 0.5em solid #313160; border-radius: 5px;" class="button-a"><!--[if mso]>&nbsp;&nbsp;&nbsp;&nbsp;<![endif]-->' + f.columButton + '<!--[if mso]>&nbsp;&nbsp;&nbsp;&nbsp;<![endif]--></a></td></tr></table></td></tr>';
						}
						if ( f.columNote !== "" ) {
							codeBody += '<tr><td style="padding: 10px 0;"></td></tr>';
						}
						// console.log('Rank');
						break;
					case 'Single-Image-Ad':
						codeBody += '<tr><td align="center" style="padding: 20px 0;"><a href="' + f.columLink + '"><img src="' + f.columImage + '" alt="' + f.columTitle + '" border="0" style="display: block; width: auto; max-width: 100%; height: auto; max-height: 100px;"></a></td></tr>';
						if ( f.columNote !== "" ) {
							codeBody += '<tr><td style="padding: 10px 0;"></td></tr>';
						}
						break;
				}
			});
		});

		// 組裝原始碼第三段，加上表尾及網頁foot的部份
		trackFoot = '<custom name="opencounter" type="tracking"/><div style="display: none; font-size: 1px; line-height: 1px; max-height: 0px; max-width: 0px; opacity: 0; overflow: hidden; mso-hide: all; font-family: sans-serif;">This email was sent by:<b>%%Member_Busname%%</b><br>%%Member_Addr%% %%Member_City%%, %%Member_State%%, %%Member_PostalCode%%, %%Member_Country%%<a href="%%profile_center_url%%" alias="Update Profile">Update Profile</a></div>';

		previewFoot = '<table align="center" width="600" class="email-container"><tr><td dir="ltr" style="padding: 40px 3.333%; width: 100%;"><table cellspacing="0" cellpadding="0" border="0" align="center" width="100%"><tr><td class="stack-column-center" style="width: 8em; font-size: 0.875rem; text-align: center;"><a href="https://www.facebook.com/opinion.cw" target="_blank" rel="noopener noreferrer" style="display: inline-block; padding: 0.5em 1em; color: #1877F2; font-size: 0.875rem; line-height: 1; background-color: #fff; border-radius: 5px;"><img src="https://topic.cw.com.tw/salesforce/assets/images/facebook.png" alt="Facebook Logo" width="20" style="display: inline-block; vertical-align: middle; margin-right: 0.5em;"><span style="display: inline-block; vertical-align: middle; line-height: 1.25rem;">追蹤獨評</span></a></td><td class="stack-column-center" style="width: 20px; height: 10px;"></td><td class="stack-column-center" style="text-align: left;"><a href="http://opinion.cw.com.tw/" target="_blank" rel="noopener noreferrer" style="padding: 0 10px 0 0; color: #000; font-size: 0.875rem; font-weight: 400; font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif; line-height: 1; mso-height-rule: exactly; text-decoration: none; border-right: 1px solid #000;">關於獨立評論@天下</a><a href="https://api.cwg.tw/web/epaper/cancel-guest?email=%%email%%&epaper=EPAPER13" target="_blank" rel="noopener noreferrer" class="unsubscribe" style="padding: 0 10px; color: #000; font-size: 0.875rem; font-weight: 400; font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif; line-height: 1; mso-height-rule: exactly; text-decoration: none; border-right: 1px solid #000;">取消訂閱</a><a href="https://opinion.cw.com.tw/blog/profile/31/article/28" target="_blank" rel="noopener noreferrer" style="padding: 0 10px; color: #000; font-size: 0.875rem; font-weight: 400; font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif; line-height: 1; mso-height-rule: exactly; text-decoration: none; border-right: 1px solid #000;">讀者投書</a><a href="tel:+886226620332" style="padding: 0 10px; color: #000; font-size: 0.875rem; font-weight: 400; font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif; line-height: 1; mso-height-rule: exactly; text-decoration: none;">讀者服務專線 (02) 2662-0332</a><div style="display: inline-block; color: #000; font-size: 0.875rem; font-weight: 400; font-family: \'Roboto\', \'Noto Sans TC\', \'思源黑體 TC\', \'思源黑體 TW\', \'思源黑體\', \'微軟正黑體\', \'繁黑體\', \'Microsoft JhengHei\', \'Lato\', \'Arial\', \'Segoe UI Emoji\', \'Segoe UI Symbol\', \'新細明體\', sans-serif; mso-height-rule: exactly;">Copyright © 2021 天下雜誌 All rights reserved.</div></td></tr></table></td></tr></table>';

		codeFoot = '</center></td></tr></table></body></html>';

		// dataLayer.push({
		//     "event": "GA-event",
		//     "eventInfo": "編碼完成",
		//     "fileID": key
		// });

		//將原始碼塞入「發信用原始碼」的框框
		$("#sourceCode textarea").val(codeHead+previewHead+codeBody+trackFoot+previewFoot+codeFoot);

		//將原始碼丟入「預覽」框，呈現組裝後的結果
		$("#previewHtml").html(previewHead+codeBody+previewFoot);
		$("#initial").html(codeHead+previewHead+codeBody+trackFoot+previewFoot+codeFoot);
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
				columChannel: f.gsx$頻道.$t,
				columChannelLink: f.gsx$頻道連結.$t,
				columAuthor: f.gsx$作者.$t,
				columButton: f.gsx$按鈕文字.$t,
				columImage: f.gsx$圖片.$t,
				columTitle: f.gsx$標題.$t,
				columText: f.gsx$內文.$t,
				columNote: f.gsx$標誌.$t
            };
        });
        document.title = data.funcColum[0].columAuthor + " - 獨立評論電子報";
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
