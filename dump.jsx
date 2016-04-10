/*
    設定ファイルを読み込んで、JSONとして返す
    ファイルが存在しない場合は、空配列を返す
*/
function setConvertTable() {
    var filename = "glyphtable.json";
    var currentDir = Folder(File(app.activeScript).path);
    var file = new File(currentDir + "/" + filename);
    
    try {
        file.open ("r");
        var text = file.read();
        // JSONに変換
        var table = eval("(" + text + ")");
    } catch(err) {
        alert("字形変換ファイルの読込時にエラーが発生しました：" + err);
        return [];
    }

    return table;
}

var table = setConvertTable();

var text = "";
for(i = 0, len = table.length; i < len; i++) {
    var dat = table[i];
    text += dat["fchr"] + "\t"+ dat["fchr"] + "\t"+ dat["fchr"]  + "\t"+ dat["fchr"] + "\n";
}
app.activeDocument.selection[0].contents  = text;
