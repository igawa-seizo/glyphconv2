var ProgressPanel = (function () {
    
    // コンストラクタ
	function ProgressPanel(title) {   
        // 最大値
        this.MaximumValue = 300 ;
        // プログレスバーの横幅
        this.ProgressBarWidth = 300 ;
        // 1単位あたりの幅
        this.Increment = this.MaximumValue / this.ProgressBarWidth;
        
        this.panel = new Window( 'window', title );
	}

    ProgressPanel.prototype.setInstance = function(maximumValue,　progressBarWidth) { 
        var panel = this.panel;
        
        with(panel){
            panel.ProgressBar = add('progressbar', [12, 12, progressBarWidth, 24], 0, maximumValue);
            panel.ProgressLabel = add( 'statictext' , [ 12 , 12 , progressBarWidth , 24 ] , '' );
            panel.MaximumValue = maximumValue ;
            panel.ProgressBarWidth = progressBarWidth ;
            panel.Increment = panel.MaximumValue / panel.ProgressBarWidth;
        }
    };

    return ProgressPanel;
}());

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

/*
    字形変換表を抽出（CID）
    table：字形変換表、jsonオブジェクト
    mode : h22, s56, trad
    flag : setConvertTableを参照、
    collection : Pro, ProN, Pr5, Pr5N, Pr6, Pr6N
*/
function makeCidTableSubset(table, mode, prefFlag, collection) {
    var list = [];
    var maxCidNumber = {
        "Pro" : 15444,
        "ProN" : 15444,
        "Pr5" : 20317,
        "Pr5N" : 20317,
        "Pr6" : 23058,
        "Pr6N" : 23058,
    };

    for(i = 0, len = table.length; i < len; i++) {
        var dat = table[i];
        var flag = 0;

        //フラグに合致するのを取り出す
        for(var tmp in dat) {
            if(parseInt(tmp) == "NaN") continue;
            var definedFlag = parseInt(tmp);
            var result = definedFlag & prefFlag;
            if(prefFlag == result) {
                flag = definedFlag;
                break;
            }
        }
        //合致しなければ次へ
        if(flag == 0) continue;
        flag = String(flag);
        
        //CIDでの変換ができないのなら無視
        if(dat[flag]["cid"] == 0 || typeof(dat[flag][mode]) != "number") continue;
        
        //CID番号がそのフォントセットより大きいときも無視
        if(dat[flag]["cid"] > maxCidNumber[collection] || dat[flag][mode] > maxCidNumber[collection]) continue;
      
        list.push( {"fchr" : dat["fchr"],
                         "src" : dat[flag]["cid"],
                         "dst" : dat[flag][mode]} );
    }
    return list;
}

/*
    字形変換表を抽出（Unicode）
    定義済みのJSONファイルで変換すると置換で時間が掛かりすぎる
    table：字形変換表、jsonオブジェクト
    mode : h22, s56, trad
    flag : setConvertTableを参照
*/
function makeUnicodeTableSubset(table, mode, prefFlag) {
    var list = [];
   
    for(i = 0, len = table.length; i < len; i++) {
        var dat = table[i];

        //フラグに合致するのを取り出す
        var flag = 0;
        var result = 0;
        for(var tmp in dat) {
            if(parseInt(tmp) == "NaN") continue;
            var definedFlag = parseInt(tmp);
            
            result = definedFlag & prefFlag;
            if(prefFlag == result) {
                flag = definedFlag;
                break;
            }
        }
        //合致しなければ次へ
        if(flag == 0) continue;
        flag = String(flag);
        //未定義（変換の必要なしの字形）は次へ
        if(dat[flag][mode] == "") continue;
        //CIDでの変換が必要なので無視
        if(typeof(dat[flag][mode]) == "number") continue;
               
        list.push({"src" : dat["fchr"],  "dst" : dat[flag][mode]});
    }
    return list;
}

var Fontset= (function () {
    var re = new RegExp("(Pro$|ProN$|Pr5$|Pr5N$|Pr6$|Pr6N$)");
    
    // コンストラクタ
	function Fontset(panel) {   
        this.openTypeFontsCnt  = 0;
        
        this.fonts = {
            "document" : {},
            "story" : {},
            "selection" : {},
        };
        this.panel = panel;
        
        this.makeFontset();
	}

    Fontset.prototype.getAppFonts = function() {
        return this.fonts["document"];
    };

    //アプリケーションのフォントを取得
    Fontset.prototype.makeFontset = function() {
        
        for (var i=0; i < app.activeDocument.fonts.length; i++){
            this.panel.ProgressBar.value++;
            
            var font =app.activeDocument.fonts[i];
            if(font.fontType.toString() != "OPENTYPE_CID") continue;
           
            var fontName = font.fontFamily + "  " + font.fontStyleName;
            if(fontName in this.fonts["document"]) continue;
            
            var result = re.exec(font.fontFamily);
            if(result != null) {
                this.fonts["document"][fontName] = font;
                this.openTypeFontsCnt++;
           }
        }
    };

    //限られた範囲でのフォントを取得
    Fontset.prototype.makeFontSubset = function(target) {
        var cnt = 0;
        var subset = {};
        
        for(var i=0; i < target.characters.length; i++) {
            var font =target.characters[i].appliedFont;
            
            if(font.fontType.toString() != "OPENTYPE_CID") continue;
            
            fontName = font.fontFamily + "  " + font.fontStyleName;
            if(fontName in subset) continue;
            
            var result = re.exec(font.fontFamily);
            if(result != null) {
                subset[fontName] = font;
                cnt++;
                
                if(cnt == this.openTypeFontsCnt) break;
            }
        }

        return subset;
    };

    return Fontset;
}());

var GlyphConverter= (function () {
    // コンストラクタ
	function GlyphConverter(pref, panel) {
        this.pref = pref;
        
        this.document  = app.activeDocument;        
        this.target = "";
        if(this.pref.target == "document") this.target = this.document;
        else if(this.pref.target == "story") this.target = this.document.selection[0].parentStory;
        else if(this.pref.target == "selection") this.target = this.document.selection[0];
        
        this.glyphTable = this.pref["glyphTable"];
        this.fontTable   = this.pref["fontTable"];
        this.progressPanel = panel;
	}

    //UNICODE内での置換を担当する
    GlyphConverter.prototype.convertText = function()  {
         //検索クエリの格納
         var fontCnt = 0;
         for(var tmp in this.fontTable) fontCnt++;
         
         if(this.pref.fontName == "使用中の全フォント" || fontCnt == 1) {
            app.findTextPreferences.appliedFont = NothingEnum.nothing;
            app.findTextPreferences.fontStyle = NothingEnum.nothing;
         } else {
             for(fontName in this.fontTable) {
                 app.findTextPreferences.appliedFont = this.fontTable[fontName];
                 app.findTextPreferences.fontStyle = this.fontTable[fontName].fontStyleName;
             }
         }
         
         var table = this.glyphTable["unicode"];
         try {
             for(var i = 0, len = table.length; i < len; i++) {
                var dat = table[i];
                //検索クエリの格納
                app.findTextPreferences.findWhat = dat["src"];
                app.changeTextPreferences.changeTo = dat["dst"];
            
                //置換の実行
                this.target.changeText();
                this.progressPanel.ProgressBar.value++;
            }
        } catch(err) {
            alert(err + "：" + dat["src"] + "→" + dat["dst"]);
        }
    };

    //CIDでの字形変換を担当する
    GlyphConverter.prototype.convertGlyph = function(fontName)  {
        //検索クエリの格納
        try {
        app.findGlyphPreferences.appliedFont = this.fontTable[fontName];
        app.findGlyphPreferences.fontStyle = this.fontTable[fontName].fontStyleName;
        
        app.changeGlyphPreferences.appliedFont = this.fontTable[fontName];
        app.changeGlyphPreferences.fontStyle = this.fontTable[fontName].fontStyleName;
        } catch(err) {
                alert(err);
                alert(fontName)
         }
        
        //変換表の取り出し
        var collection = getCorrespondence(fontName);
        var table = this.pref["glyphTable"]["cid"][collection];
        try {
            for(var i = 0, len = table.length; i < len; i++) {
                var dat = table[i];
                app.findGlyphPreferences.glyphID = dat["src"];
                app.changeGlyphPreferences.glyphID = dat["dst"];
                
                this.target.changeGlyph();            
                this.progressPanel.ProgressBar.value++;
            }
        } catch(err) {
            alert(err + ":" + dat["fchr"] + "（" + fontName +"）、" + dat["src"] + "→" + dat["dst"]);
        }
    };    

    GlyphConverter.prototype.convert = function () {
        //UNICODE内での変換
       this.convertText();

        //CID変換
        for(var fontName in this.fontTable) {
            this.convertGlyph(fontName);
        }
    }

    function getCorrespondence(fontName) {
        //変換表の取り出し
        var re = new RegExp("(ProN|Pro|Pr5N|Pr5|Pr6N|Pr6)");
        var result = re.exec(fontName);
        var correspondence= {
            "Pro" : "Pro",
            "ProN" : "ProN",
            "Pr5" : "Pr6",
            "Pr5N" : "Pr6N",
            "Pr6" : "Pr6",
            "Pr6N" : "Pr6N",
        };
        return correspondence[result[0]];
    }

    return GlyphConverter;
}());


//設定ダイアログ
var SettingDialog = (function ()  {
    var instance;
    var dialog;
    
	// コンストラクタ
	function SettingDialog(appFonts) {
	     instance = this;
         this.appFonts = appFonts;
         this.font;
         this.range;
         
         this.rangeSet = {
            0 : "document",
            1 : "story",
            2 : "selection",
         };
     
         this.glyphSet = {
            0 : "h22_fude",
            1 : "h22",
            2 : "trad",
         };
        
		return instance;
	}

    //ダイアログ設定
    SettingDialog.prototype.setWindow = function()  {
       dialog = app.dialogs.add({name:"字形変換オプション", canCancel:true});
        
        with(dialog){
            with(dialogColumns.add()) {                
                with(borderPanels.add()) {
                    with(dialogColumns.add()) {  
                        this.setFontMenu(dialogRows.add());
                        with(dialogRows.add()) {　staticTexts.add({staticLabel:""}); }
                        this.setGlyphMenu(dialogRows.add());
                        with(dialogRows.add()) {　staticTexts.add({staticLabel:""});}
                        this.setRangeMenu(dialogRows.add());
                    }
                }
            }
        }
    };

    //フォントメニュー
    SettingDialog.prototype.setFontMenu= function(panel)  {
        var list = [];
        list.push("使用中の全フォント")
        for(font in this.appFonts) {
            list.push(font);
        }
        with(panel) {
            staticTexts.add({staticLabel:"フォント："});
            this.font = dropdowns.add({stringList: list, selectedIndex:0, minWidth:200});
        }
    };

    //字形変換メニュー
    SettingDialog.prototype.setGlyphMenu= function(panel)  {
        with(panel) {
            staticTexts.add({staticLabel:"字形の変換："});
            
            this.glyph = radiobuttonGroups.add();
            with(this.glyph){
                radiobuttonControls.add({staticLabel:"常用漢字表（平成22年）／筆押さえあり", checkedState:true});
                radiobuttonControls.add({staticLabel:"常用漢字表（平成22年）／筆押さえなし"});
                radiobuttonControls.add({staticLabel:"いわゆる康熙字典体（旧字体）"});
            }
        }
    };
 
     //検索範囲メニュー
     SettingDialog.prototype.setRangeMenu= function(panel)  {
         var doc = app.activeDocument;
         var list = [];
         list.push("ドキュメント");
        if(doc.selection.length > 0) { 
            list.push("ストーリー");
            if(doc.selection[0].characters.length > 0) {
                list.push("選択範囲");
            }
        }
         
         with(panel) {
            staticTexts.add({staticLabel:"検索と置換："});
            this.range = dropdowns.add({stringList: list, selectedIndex:0, minWidth:200});
        }
     };

    //ダイアログ描写
     SettingDialog.prototype.show= function()  {
          if(dialog.show() == true) {
              return 0;
          } else {
              return 1;
          }
     };
 
    //ダイアログ破壊
     SettingDialog.prototype.dispose= function()  {
          dialog = null;
     };

    //フォントの選択値取得
	SettingDialog.prototype.getFont= function()  {
           return this.font.stringList[this.font.selectedIndex];     
     };
 
    //対象の選択値取得
    SettingDialog.prototype.getRange= function()  {  
           return this.rangeSet[this.range.selectedIndex];
     };
 
     //字形の選択値取得
    SettingDialog.prototype.getGlyphPref= function()  {
           return this.glyphSet[this.glyph.selectedButton];
     };
    
	return SettingDialog;
}());

////////////////////////////////////////////////////////////////////////////

function main() {
    //初期化処理
    var convertTable = setConvertTable();
    if(convertTable.length == 0) return;
    
    //OpenTypeフォントの取得
    var panelObj = new ProgressPanel("フォント取得中");
    panelObj.setInstance(app.activeDocument.fonts.length, 400);
    panelObj.panel.ProgressLabel.text  = "フォントを取得中です。完了までしばらくお待ちください……" ;
    panelObj.panel.show();

    var fs = new Fontset(panelObj.panel);

    panelObj.panel.close();

    var dialog = new SettingDialog(fs.getAppFonts());
    //字形変換ダイアログの設定
    while(1) {
         dialog.setWindow();
         if(dialog.show() != 0) {
             dialog.dispose();
             break;
         }
         
         var pref = { 
              "fontName" : dialog.getFont(),
              "target" : dialog.getRange(),
              "mode" : dialog.getGlyphPref(),
              "fontTable" : {},
              "glyphTable" : {},
         };
     
        //設定に応じたる字形変換表を作成する
        pref["glyphTable"]  = {
            "unicode" : makeUnicodeTableSubset(convertTable, pref["mode"], 3),
            "cid" : {
                "Pro"   : makeCidTableSubset(convertTable, pref["mode"], 1,   "Pro"),
                "ProN" : makeCidTableSubset(convertTable, pref["mode"], 2, "ProN"),
                "Pr6"   : makeCidTableSubset(convertTable, pref["mode"], 1,   "Pr6"),
                "Pr6N" : makeCidTableSubset(convertTable, pref["mode"], 2, "Pr6N"),
            },
        };
        
        //フォント一覧表の設定
        if(pref["fontName"] == "使用中の全フォント") {
            pref["fontTable"] = fs.fonts["document"];
        } else {
            pref["fontTable"][pref["fontName"]] = fs.fonts["document"][pref["fontName"]];
        }

        //変換工数の計算
        var unicodeTableSize = pref["glyphTable"]["unicode"].length;
        var cidTableSize = pref["glyphTable"]["cid"]["Pr6N"].length;

        var cnt = 0;
        if(pref["fontName"] == "使用中の全フォント") {
            cnt = unicodeTableSize + app.activeDocument.fonts.length * cidTableSize;
        } else {
            cnt = unicodeTableSize + cidTableSize;
        }
        
        //進捗ダイアログの設定
        var panelObj = new ProgressPanel("字形変換中");
        panelObj.setInstance(cnt, 400);
        panelObj.panel.ProgressLabel.text  = "字形変換中です。完了までしばらくお待ちください……" ;
        panelObj.panel.show();

        //タイマーを設定
        var startTime = new Date();

        //コントローラを作成
        var ctrl= new GlyphConverter(pref, panelObj.panel);    
        ctrl.convert();
        var endTime = new Date();
        
        alert("変換が終了しました。所要時間：" + (endTime - startTime) / 1000 + "秒");
        panelObj.panel.close();  
    }
}

//メイン関数呼び出し
main();

