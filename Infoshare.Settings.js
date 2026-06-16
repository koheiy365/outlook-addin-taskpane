
// ToDo: Configuration.xmlと関連する設定情報の保持ロジック全体の検討

//コンテキスト設定：iVxContext
iSettings.iVxContext = {};
iSettings.iVxContext.IsOnline = true;

// カスタムフォーム設定:iSPListFormExtender
iSettings.iSPListFormExtender = {};
iSettings.iSPListFormExtender.AdditionalExtendForms =
    [
        // 拡張を行うフォーム名を以下のように定義する
        //"TestList01.aspx",
        //"TestList02.aspx",
        //"TestList03.aspx"
    ];

// モダンUI用コンテキスト管理用
iSettings.MasterSiteContext = {};
iSettings.MasterSiteContext.ServerRelativePath = "/sites/adminSiteName";