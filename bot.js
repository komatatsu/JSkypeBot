var skype = new ActiveXObject('Skype4COM.Skype');
WScript.ConnectObject(skype, 'Skype_');
skype.Attach();

//���[�v���񂵂đҋ@
while(true){
    WScript.Sleep(1000);
}

/**
 * Skype�̃��b�Z�[�W���󂯂�
 */
function Skype_MessageStatus(msg, status){
    var user = msg.FromDisplayName;    //���M�Җ�
    var body = msg.Body;               //���b�Z�[�W
    var chat = msg.Chat;               //�`���b�g���

    if(status == skype.Convert.TextToChatMessageStatus('RECEIVED')
        || status == skype.Convert.TextToChatMessageStatus('SENDING')){
        //�w���v�\��
        if (body.indexOf("@help") === 0) {
            chat.SendMessage("�ق��ق�");
        }
        //Redmine�̃��r���[�҂��`�P�b�g�ꗗ�\��
        if (body.indexOf("@review") === 0) {
            getTicketList4Review(chat);
        }
    }
}

/**
 * ���r���[�҂��̃`�P�b�g����Atom�t�B�[�h����擾
 */
function getTicketList4Review(chat) {
    chat.SendMessage("������Ƒ҂��Ă�");
    //FIXME: �`�P�b�g�ꗗ��URL
    var url = "http://hoge.com/redmine/projects/hogehoge/issues.atom?";
    var xml = new ActiveXObject("Microsoft.XMLDOM");
    xml.async = false;
    xml.load(url);

    setNode(xml.documentElement, chat);
}

/**
 * Atom�t�B�[�h�̉�͏���
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
