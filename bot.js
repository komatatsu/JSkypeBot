var skype = new ActiveXObject('Skype4COM.Skype');
WScript.ConnectObject(skype, 'Skype_');
skype.Attach();

//ループを回して待機
while(true){
    WScript.Sleep(1000);
}

/**
 * Skypeのメッセージを受ける
 */
function Skype_MessageStatus(msg, status){
    var user = msg.FromDisplayName;    //送信者名
    var body = msg.Body;               //メッセージ
    var chat = msg.Chat;               //チャット情報

    if(status == skype.Convert.TextToChatMessageStatus('RECEIVED')
        || status == skype.Convert.TextToChatMessageStatus('SENDING')){
        //ヘルプ表示
        if (body.indexOf("@help") === 0) {
            chat.SendMessage("ほげほげ");
        }
        //Redmineのレビュー待ちチケット一覧表示
        if (body.indexOf("@review") === 0) {
            getTicketList4Review(chat);
        }
    }
}

/**
 * レビュー待ちのチケット情報をAtomフィードから取得
 */
function getTicketList4Review(chat) {
    chat.SendMessage("ちょっと待ってね");
    //FIXME: チケット一覧のURL
    var url = "http://hoge.com/redmine/projects/hogehoge/issues.atom?";
    var xml = new ActiveXObject("Microsoft.XMLDOM");
    xml.async = false;
    xml.load(url);

    setNode(xml.documentElement, chat);
}

/**
 * Atomフィードの解析処理
 */
function setNode(node, chat) {
    if(node == null) {
        return;
    }

    if(node.nodeName == "entry"){
        for (var i = 0, len = node.childNodes.length; i < len; i++) {
            var child = node.childNodes.item(i);
            if (child.nodeName == "title") {
                chat.SendMessage(child.text);
            }
        }
    }
    setNode(node.firstChild, chat);
    setNode(node.nextSibling, chat);
}
