// =======================================================================
//
//  TaskPane.js
//
//                         (C) 2010-17 Infoshare Inc. All rights reserved
// =======================================================================

// -----------------------------------------------------------------------
//  Register namespace
// -----------------------------------------------------------------------
ISNamespace.Register('Infoshare.GroupPicker');


// -----------------------------------------------------------------------
//  OutlookExtension
// -----------------------------------------------------------------------
Infoshare.GroupPicker.OutlookExtension = function Infoshare$GroupPicker$OutlookExtension() { }
GPOutlookExtension = Infoshare.GroupPicker.OutlookExtension;

GPOutlookExtension.PlatformType = {
    Outlook: "Outlook",
    OWA: "OWA"
};
GPOutlookExtension.FormType = {
    Read: "Read",
    Edit: "Edit"
};
GPOutlookExtension.ApplicationType = {
    Mail: "Mail",
    Appointment: "Appointment"
};

GPOutlookExtension.Platform = null;
GPOutlookExtension.Form = null;
GPOutlookExtension.Application = null;

GPOutlookExtension.IsVstoAddin = false;

// ToDo: Resxファイルに切り出す
GPOutlookExtension.ErrorMessage = {
    EnUS : {
        LicenseInValid     : "The product license is registered incorrectly or expired. Please contact your administrator.",
        ConfiguratonProblem: "There is a problem with the configuration. Please contact your administrator.",
        NotSuportBrowser   : "This browser is not supported.",
        DisabledUser       : "You are prohibited to use this application. Please contact your administrator."
    },
    JaJp : {
        LicenseInValid     : "製品ライセンスが正しく登録されていないか有効期限が切れています。管理者にお問い合わせください。",
        ConfiguratonProblem: "構成に問題があります。管理者にお問い合わせください。",
        NotSuportBrowser   : "サポートされていないブラウザです。",
        DisabledUser       : "このアプリの使用を禁止されています。管理者にお問い合わせください。"
    }
};

// コンテキスト判断用にInfoCallのWindowをキャッシュする
GPOutlookExtension.InfoCallWindow = null;

GPOutlookExtension.InitializeInfoCallPane = function()
{
    // ライセンスや機能の不備がある場合には管理者に問い合わせる旨を通知する
    var contentDiv = document.getElementById("content-main");
    var rerourceObject = GPOutlookExtension.ErrorMessage.EnUS;
    /*
    if (_spPageContextInfo)
    {
        rerourceObject = _spPageContextInfo.currentCultureName == "ja-JP" ? GPOutlookExtension.ErrorMessage.JaJp
                                                                          : GPOutlookExtension.ErrorMessage.EnUS;
    }
    else
    {
        // ToDo: Outlookの言語設定の取得
    }    

    // 構成不備
    if (typeof(iVxContext) == "undefined" || // InfoShare.jsが読み込まれていない
        typeof(GPContext) == "undefined"  || // GroupPicker.jsが読み込まれていない
        typeof (GPSettings.PhoneBookFeature) == "undefined" ||
        !GPSettings.PhoneBookFeature) // 電話帳機能が有効になっていない
    {
        GPOutlookExtension.DisplayAlertMessage(rerourceObject.ConfiguratonProblem);
        return;
    }

    // ブラウザ確認
    // Outlookでのみ、Edgeでも動作するようにブラウザの判断を変更。
    // 以下の条件で、Outlook が利用するレンダリングエンジンが、Trident（Internet Explorer）からEdgeHTML（Edge）に切り替わる。
    // Windows 10 ver. >= 1903 / Office 365 ver >= 16.0.11629
    // https://docs.microsoft.com/en-us/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins
    if (!GPContext.IsSupportedBrowser() && !GPContext.IsOfficeAddInSupportedBrowser())
    {
        GPOutlookExtension.DisplayAlertMessage(rerourceObject.NotSuportBrowser);
        return;
    }

    // OWAはEdgeを許可しない
    try
    {
        if (iString.Equals(Office.context.platform, "OfficeOnline", true) && iBrowser.IsEdgeLegacy())
        {
            GPOutlookExtension.DisplayAlertMessage(rerourceObject.NotSuportBrowser);
            return;
        }
    }
    catch (e) { }

    // コンテキストの初期化
    GPContext.InitializeOption.EnsureResource = false;
    GPContext.InitializeOption.FileLoad = false;
    GPContext.InitializeOption.SiteConfigurationInitialize = false;
    GPContext.Initialize(function ()
    {
        var license = GPContext.InfoCallLicense;
            
        // ライセンス確認
        if (!license ||
            !license.Valid)
        {
            // ライセンス切れを通知する
            GPOutlookExtension.DisplayAlertMessage(rerourceObject.LicenseInValid);
            return;
        }

        // ユーザーごとの設定確認
        var disabled = false;
        GPViewContext.InitializeOption.FileLoad = false;
        GPViewContext.Initialize(function ()
        {
            var dataSources = GPViewContext.DataSources.GetValues();
            var defaultDataSource = dataSources && dataSources.length > 0 ? dataSources[0]
                                                                          : null;
            var currentUserEntity = defaultDataSource ? defaultDataSource.CurrentUserEntity
                                                      : null;
            if (currentUserEntity)
            {
                var disabledAddin = currentUserEntity.Properties["DisabledOutlookAddin"];
                disabled = ISBoolean.TryParse(disabledAddin, false)
            }
            if (disabled)
            {
                // 禁止されている旨を通知する
                GPOutlookExtension.DisplayAlertMessage(rerourceObject.DisabledUser);
                return;
            }

            // コンテンツを表示する
            contentDiv.style.display = "";

            // アドインの初期化
            GPOutlookExtension.InitializeInfoCallPaneInternal();
        });
    });
    */
    // コンテンツを表示する
    contentDiv.style.display = "";

    // アドインの初期化
    GPOutlookExtension.InitializeInfoCallPaneInternal();
};

GPOutlookExtension.DisplayAlertMessage = function(message)
{
    var contentDiv = document.getElementById("content-main");
    contentDiv.innerHTML = message;
    contentDiv.style.display = "";
}

GPOutlookExtension.InitializeInfoCallPaneInternal = function()
{
    // コンテキストの確認
    GPOutlookExtension.EnsureContext();

    // コンテキストに応じてフォームを拡張する

    // InfoCallページを開く
    setTimeout(function ()
    {
        GPOutlookExtension.OpenInfoCallLink();
    }, 1024);

    // ペインのクローズイベントの設定（OutlookはOffice.js内の実装で十分なのでOWAのみ対応する）
    if (GPOutlookExtension.Platform == GPOutlookExtension.PlatformType.OWA)
    {
        window.onunload = function()
        {
            GPOutlookExtension.Close();
        }
    }
};

GPOutlookExtension.EnsureContext = function()
{
    // VSTO
    if (GPOutlookExtension.IsVstoAddin)
    {
        GPOutlookExtension.EnsureContextForVstoAddin();
    }
    // Office
    else
    {
        GPOutlookExtension.EnsureContextForOfficeAddin();
    }
};

GPOutlookExtension.EnsureContextForOfficeAddin = function()
{
    // プラットフォーム: 基本的にはOffice.context.uiの有無で判断する
    // OnlineのOWAのみOffice.context.uiを保持しているため、platformプロパティを確認する
    var platform = Office.context.platform;
    if (platform)
    {
        if (ISString.Equals(platform, "OfficeOnline", true))
        {
            GPOutlookExtension.Platform = GPOutlookExtension.PlatformType.OWA;
        }
        else
        {
            GPOutlookExtension.Platform = GPOutlookExtension.PlatformType.Outlook;
        }    
    }
    else
    {
        if (Office.context.ui)
        {
            GPOutlookExtension.Platform = GPOutlookExtension.PlatformType.Outlook;
        }
        else
        {
            GPOutlookExtension.Platform = GPOutlookExtension.PlatformType.OWA;
        }        
    }
        
    // アプリケーション種別
    var item = Office.context.mailbox.item;
    var itemType = item.itemType;
    if (ISString.Equals(itemType, "Message", true))
    {
        GPOutlookExtension.Application = GPOutlookExtension.ApplicationType.Mail;
    }
    else if (ISString.Equals(itemType, "Appointment", true))
    {
        GPOutlookExtension.Application = GPOutlookExtension.ApplicationType.Appointment;
    }
};

GPOutlookExtension.EnsureContextForVstoAddin = function()
{
    // プラットフォーム: OWAは存在しないので固定値
    GPOutlookExtension.Platform = GPOutlookExtension.PlatformType.Outlook;
        
    // アプリケーション種別：クエリストリングから判断する
    var application = ISQueryString.GetValue("Application");
    if (ISString.Equals(application, "Mail", true))
    {
        GPOutlookExtension.Application = GPOutlookExtension.ApplicationType.Mail;
    }
    else if (ISString.Equals(application, "Appointment", true))
    {
        GPOutlookExtension.Application = GPOutlookExtension.ApplicationType.Appointment;
    }
};

GPOutlookExtension.OpenInfoCallLink = function()
{
    // VSTO
    if (GPOutlookExtension.IsVstoAddin)
    {
        GPOutlookExtension.OpenInfoCallLinkForVstoAddin();
    }
    // Office
    else
    {
        GPOutlookExtension.OpenInfoCallLinkForOfficeAddin();
    }
};

GPOutlookExtension.OpenInfoCallLinkForOfficeAddin = function()
{
    // 既に開かれている場合
    if (GPOutlookExtension.InfoCallWindow)
    {
        return;
    }
    var href = location.href;
    var search = location.search;
    var hrefWitoutSerch = href.replace(search, "");
    var hash = location.hash;
    var hrefWithoutHash = hrefWitoutSerch.replace(hash, "");
    //var infoCallPagePath = hrefWithoutHash + "/../../../../OutlookAddinPicker.aspx";
    var infoCallPagePath = "https://infosharedev.sharepoint.com/sites/Dev_KoheiY15/InfoshareProductLibrary/Sources/Infoshare.InfoCall/Extensions/Outlook/Sources/OutlookAddinPicker.aspx;

    var queryString = "?"
                    + "Platform="
                    + GPOutlookExtension.Platform
                    + "&Form="
                    + GPOutlookExtension.Form
                    + "&Application="
                    + GPOutlookExtension.Application;
    var url = infoCallPagePath + queryString;

    // Outlookクライアントはダイアログで開く
    // ToDo：2013、2016のバージョン差異によってサイズ比較の対象が異なるので、ロジックを分岐させる
    // 2013：ディスプレイのサイズ
    // 2016：メールフォームのサイズ
    if (GPOutlookExtension.Platform == GPOutlookExtension.PlatformType.Outlook)
    {
        var width  = 50;
        var height = 50;
        if (GPSettings.GPOutlookExtension && GPSettings.GPOutlookExtension.OutlookDialogSize)
        {
            width  = GPSettings.GPOutlookExtension.OutlookDialogSize.Width;
            height = GPSettings.GPOutlookExtension.OutlookDialogSize.Height;
        }

        Office.context.ui.displayDialogAsync(url, { height: height, width: width }, function (asyncResult)
        {
            if (asyncResult.status == "succeeded")
            {
                GPOutlookExtension.InfoCallWindow = asyncResult.value;

                // イベント追加
                GPOutlookExtension.InfoCallWindow.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, GPOutlookExtension.SetEntities);
                GPOutlookExtension.InfoCallWindow.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogEventReceived, GPOutlookExtension.DialogEventReceived);
            }
        });
    }
    // OWAの場合はnewWindowで開く
    else
    {
        var width  = 800;
        var height = 600;
        if (GPSettings.GPOutlookExtension && GPSettings.GPOutlookExtension.OWADialogSize)
        {
            width  = GPSettings.GPOutlookExtension.OWADialogSize.Width;
            height = GPSettings.GPOutlookExtension.OWADialogSize.Height;
        }
        
        // 複数フォームが開かれた場合に、固有の名前windowに設定する必要があるため、Guidを設定
        var option = ISString.Format("width={0}, height={1}, resizable=yes", width, height);
        var windowName = ISGuid.New();
        var newWindow = window.open(url, windowName, option);
                
        GPOutlookExtension.InfoCallWindow = newWindow;
        if (newWindow)
        {
            // リダイレクトでイベントがキャッシュされないため
            // 以下のメソッドをInfoshare.GroupPicker.js（Finder.js）内で登録する
            //newWindow.onunload = function()
            //{
            //    // クローズ処理
            //    GPOutlookExtension.Close();
                
            //}; 
        }
    }
};

GPOutlookExtension.OpenInfoCallLinkForVstoAddin = function()
{
    window.external.OpenDialogForm();
};

GPOutlookExtension.SetEntities = function(args)
{
    // Mail
    var entitiesString = args.message;
    if (!ISString.IsNullOrEmpty(entitiesString))
    {
        if (GPOutlookExtension.Application == GPOutlookExtension.ApplicationType.Mail)
        {
        
            var toPrefix = "[ToItems]" ;
            var ccPrefix = "[CCItems]";
            var bccPrefix = "[BCCItems]";
            var toIndex = entitiesString.indexOf(toPrefix);
            var ccIndex = entitiesString.indexOf(ccPrefix);
            var bccIndex = entitiesString.indexOf(bccPrefix);


            // To
            var toEntitiesString = entitiesString.substring(toIndex + toPrefix.length, ccIndex);
            var ccEntitiesString = entitiesString.substring(ccIndex + ccPrefix.length, bccIndex);
            var bccEntitiesString = entitiesString.substring(bccIndex + bccPrefix.length, entitiesString.length);

            // stringのエンティティから分割
            var toArray = ISString.Split(toEntitiesString, ";");
            var ccArray = ISString.Split(ccEntitiesString, ";");
            var bccArray = ISString.Split(bccEntitiesString, ";");

            // To
            Office.context.mailbox.item.to.addAsync(toArray);
           
            // CC
            Office.context.mailbox.item.cc.addAsync(ccArray);

            // BCC
            Office.context.mailbox.item.bcc.addAsync(bccArray);


        }
        // 予定表
        else
        {
            var entitiesString = args.message;
            var attendeeArray = ISString.Split(entitiesString, ";");

            // 必須参加者に追加
            Office.context.mailbox.item.requiredAttendees.addAsync(attendeeArray);
        }
    }

    // クローズ処理
    GPOutlookExtension.Close();
};

GPOutlookExtension.DialogEventReceived = function(args)
{
    // クローズイベントのキャッチ
    switch (args.error) {
        case 12002:
            //showNotification("Cannot load URL, 404 not found?");
            break;
        case 12003:
            //showNotification("Invalid URL Syntax");
            break;
        case 12004:
            //showNotification("Domain not in AppDomain list");
            var test = 1;
            break;
        case 12005:
            //showNotification("HTTPS Required");
            break;
        case 12006:            
            // クローズ処理           
            GPOutlookExtension.Close();
            break;
    }
};

GPOutlookExtension.SetMailItems = function(mailItems)
{
    // To
    Office.context.mailbox.item.to.addAsync(GPOutlookExtension.GetMailAddressArray(mailItems.To));
           
    // CC
    Office.context.mailbox.item.cc.addAsync(GPOutlookExtension.GetMailAddressArray(mailItems.CC));

    // BCC
    Office.context.mailbox.item.bcc.addAsync(GPOutlookExtension.GetMailAddressArray(mailItems.BCC));

    // クローズ処理
    GPOutlookExtension.Close();
};


GPOutlookExtension.SetAppointmentItems = function(appointmentItems)
{
    // requiredAttendees
    Office.context.mailbox.item.requiredAttendees.addAsync(GPOutlookExtension.GetMailAddressArray(appointmentItems));

    // クローズ処理
    GPOutlookExtension.Close();
};

GPOutlookExtension.GetMailAddressArray = function(items)
{
    var array = [];

    for (var i = 0; i < items.length; i++)
    {
        array.push(items[i].Properties.Mail);
    }

    return array;
};

GPOutlookExtension.Close = function()
{
    try
    {
        // Window・ダイアログクローズ
        if (GPOutlookExtension.InfoCallWindow)
        {
            GPOutlookExtension.InfoCallWindow.close();
        }
    }
    catch(e){}

    // Outlook2013ではクローズイベントにてタスクペインを閉じることができない為、共通的にペインのクローズは行わない
    // ペインクローズ
    //if (Office.context.ui &&
    //    Office.context.ui.closeContainer)
    //{
    //    Office.context.ui.closeContainer();
    //}
    //else if (Office.context.mailbox.CloseApp)
    //{
    //        Office.context.mailbox.CloseApp();
    //}    

    // Windowキャッシュをクリア
    GPOutlookExtension.InfoCallWindow = null;
};

// OfficeAddin初期化メソッド
GPOutlookExtension.InitializeOfficeAddin = function()
{
    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason)
    {
        $(document).ready(function ()
        {
            app.initialize();

            GPOutlookExtension.InitializeInfoCallPane();
        });
    };
};

// VSTOAddin初期化メソッド
GPOutlookExtension.InitializeVstoAddin = function()
{
    $(document).ready(function ()
    {
        GPOutlookExtension.InitializeInfoCallPane();
    });
};

// Onloadに追加
(function ()
{
    // クエリを判断
    var isVstoAddin = ISQueryString.GetValue("IsVstoAddin");
    if (ISBoolean.TryParse(isVstoAddin, false))
    {
        GPOutlookExtension.IsVstoAddin = true;
        GPOutlookExtension.InitializeVstoAddin();
    }
    else
    {
        GPOutlookExtension.InitializeOfficeAddin();
    }
})();