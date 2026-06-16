
// 関数拡張設定：GPFunctionExtender
GPSettings.GPFunctionExtender = {};
GPSettings.GPFunctionExtender.CanExtend = true;
GPSettings.GPFunctionExtender.FunctionExtenderClassNames = [
    "Infoshare.SharePoint.SharingFunctionExtender",
    "Infoshare.GroupPicker.SharingFunctionExtender",
    "Infoshare.GroupPicker.UserControlExtender"
];

// Outlookアドインの設定
// OWA：ピクセル値で指定する
GPSettings.GPOutlookExtension = {};
GPSettings.GPOutlookExtension.OWADialogSize = {
    Width : 800,
    Height: 600
};
// Outlook：ウィンドウサイズに対するパーセンテージで指定する
GPSettings.GPOutlookExtension.OutlookDialogSize = {
    Width : 50,
    Height: 50
};