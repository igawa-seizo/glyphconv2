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
        alert("設定ファイルが存在しません");
        return [];
    }

    return table;
}

/*
    字形変換表を抽出
    mode : h22, s56, trad
    flag : setConvertTableを参照、
    collection : Pro, ProN, Pr5, Pr5N, Pr6, Pr6N
*/

function makeCidTableSubset(mode, prefFlag, collection) {
    var list = [];
    var maxCidNumber = {
        "Pro" : 15444,
        "ProN" : 15444,
        "Pr5" : 20317,
        "Pr5N" : 20317,
        "Pr6" : 23058,
        "Pr6N" : 23058,
    };

    var len = convertTable.length;
    for(var i = 0; i < len; i++) {
        var dat = convertTable[i];
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

function makeUnicodeTableSubset(mode, prefFlag) {
    var list = [];
   
    var len = convertTable.length;
    for(var i = 0; i < len; i++) {
        var dat = convertTable[i];

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
        
        this.appFonts = {};
        this.panel = panel;
        
        this.makeFontset();
	}

    Fontset.prototype.getAppFonts = function() {
        return this.appFonts;
    };

    //アプリケーションのフォントを取得
    Fontset.prototype.makeFontset = function() {
        
        for (var i=0; i < app.activeDocument.fonts.length; i++){
            this.panel.ProgressBar.value++;
            
            var font =app.activeDocument.fonts[i];
            if(font.fontType.toString() != "OPENTYPE_CID") continue;
           
            var fontName = font.fontFamily + "  " + font.fontStyleName;
            if(fontName in this.appFonts) continue;
            
            var result = re.exec(font.fontFamily);
            if(result != null) {
                this.appFonts[fontName] = font;
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
        this.story = this.document.selection[0].parentStory;
        this.selection = this.document.selection;
        
        this.target = "";
        if(this.pref.target == "document") this.target = this.document;
        else if(this.pref.target == "story") this.target = this.story;
        else if(this.pref.target == "selection") this.target = this.selection[0];
        
        this.glyphTable = this.pref["glyphTable"];
        this.fontTable   = this.pref["fontTable"];
        this.progressPanel = panel;
	}

    //UNICODE内での置換を担当する
    GlyphConverter.prototype.convertText = function()  {
         //フォントによって選ぶ
         if(this.pref.fontSet == "使用中の全フォント") {
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
         /*for (var i=0; i<this.document.pages.length; i++){
                var pageObj = this.document.pages[i];
                for (var j=0; j<pageObj.textFrames.length; j++) {
                     pageObj.textFrames[j].parentStory.changeText();
               }
         }*/
    };

    //字形変換
    GlyphConverter.prototype.convertGlyph = function(fontName)  {
        //検索クエリの格納
        app.findGlyphPreferences.appliedFont = this.fontTable[fontName];
        app.findGlyphPreferences.fontStyle = this.fontTable[fontName].fontStyleName;
        app.changeGlyphPreferences.appliedFont = this.fontTable[fontName];
        app.changeGlyphPreferences.fontStyle = this.fontTable[fontName].fontStyleName;
        
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
        var collection = correspondence[result[0]];
        
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
         if(this.pref.target == "selection" && this.selection.length == 0) {
             return;
        }

        //UNICODE内での変換
       this.convertText();

        //CID変換
        for(var fontName in this.fontTable) {
            this.convertGlyph(fontName);
        }
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
            0 : "h22_hige",
            1 : "h22_none",
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
                    staticTexts.add({staticLabel:"フォント："});

                    var list = [];
                    list.push("使用中の全フォント")
                    for(font in this.appFonts) {
                        list.push(font);
                    }
                    this.font = dropdowns.add({stringList: list, selectedIndex:0, minWidth:200});
                    
                }

                with(borderPanels.add()){
                    staticTexts.add({staticLabel:"字形変換設定："});
                    this.glyph = radiobuttonGroups.add();
                    with(this.glyph){
                        radiobuttonControls.add({staticLabel:"常用漢字表（平成22年）／筆押さえあり", checkedState:true});
                        radiobuttonControls.add({staticLabel:"常用漢字表（平成22年）／筆押さえなし"});
                        radiobuttonControls.add({staticLabel:"所謂康熙字典体（旧字体）"});
                    }
                }

                with(borderPanels.add()){
                    staticTexts.add({staticLabel:"検索と置換範囲："});
                    this.range = radiobuttonGroups.add();
                    with(this.range){
                        radiobuttonControls.add({staticLabel:"ドキュメント", checkedState:true});
                        radiobuttonControls.add({staticLabel:"ストーリー"});
                        radiobuttonControls.add({staticLabel:"選択範囲"});
                    }
                }
            }
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
          this.instance = null;
     };

    //フォントの選択値取得
	SettingDialog.prototype.getFont= function()  {
           return this.font.stringList[this.font.selectedIndex];     
     };
 
    //対象の選択値取得
    SettingDialog.prototype.getRange= function()  {
           return this.rangeSet[this.range.selectedButton];
     };
 
     //字形の選択値取得
    SettingDialog.prototype.getGlyphPref= function()  {
           return this.glyphSet[this.glyph.selectedButton];
     };
    
	return SettingDialog;
}());

////////////////////////////////////////////////////////////////////////////
//初期化処理
var convertTable = setConvertTable();

//OpenTypeフォントの取得
var panelObj = new ProgressPanel("フォント取得中");
panelObj.setInstance(app.activeDocument.fonts.length, 400);
panelObj.panel.ProgressLabel.text  = "フォントを取得中です。完了までしばらくお待ちください……" ;
panelObj.panel.show();

var fs = new Fontset(panelObj.panel);
var appFonts = fs.getAppFonts();
var storyFonts = fs.makeFontSubset(app.activeDocument.selection[0].parentStory);
var selectionFonts = fs.makeFontSubset(app.activeDocument.selection[0]);

panelObj.panel.close();

var dialog = new SettingDialog(fs.getAppFonts());
while(1) {
    //字形変換ダイアログの設定
     dialog.setWindow();
     if(dialog.show() != 0) {
         dialog.dispose();
         break;
     }
     
     var pref = { 
          "fontSet" : dialog.getFont(),
          "target" : dialog.getRange(),
          "mode" : "trad",
          "hige" : 1,
          "fontTable" : {},
          "glyphTable" : {
                "unicode" : [],
                "cid" : {
                    "Pro" : [],
                    "ProN" : [],
                    "Pr6" : [],
                    "Pr6N" : [],
                },
          },
     };

    //字形変換設定の取得
    var glyphPref = dialog.getGlyphPref(); 
    if(glyphPref == "h22_hige") {
        pref["mode"] = "h22_fude";
    } else if(glyphPref == "h22_none") {
        pref["mode"] = "h22";
    } else if(glyphPref == "trad") {
        pref["mode"] = "trad";
    } 
    
    //フォント一覧表の設定
    if(dialog.getFont() == "使用中の全フォント") {
        if(pref["target"] == "document") pref["fontTable"] = appFonts;
        else if(pref["target"] == "story") pref["fontTable"] = storyFonts;
        else if(pref["target"] == "selection") pref["fontTable"] = selectionFonts;
    } else {
        var fontName = dialog.getFont();
        pref["fontTable"][fontName] = appFonts[fontName];
    }
    
    //各フォント（Pro、ProN、Pr6、Pr6N）に応じたる変換表を作成
    var panelObj = new ProgressPanel("変換テーブルの作成中");
    panelObj.setInstance(5, 400);
    panelObj.panel.ProgressLabel.text  = "字形変換テーブルの作成中です。完了までしばらくお待ちください……" ;
    panelObj.panel.show();
    
    pref["glyphTable"]["unicode"] = makeUnicodeTableSubset(pref["mode"], 15);
    panelObj.panel.ProgressBar.value++;
    
    pref["glyphTable"]["cid"]["Pro"]    = makeCidTableSubset(pref["mode"], 7, "Pro");
    panelObj.panel.ProgressBar.value++;
    pref["glyphTable"]["cid"]["ProN"] = makeCidTableSubset(pref["mode"], 11, "ProN");
    panelObj.panel.ProgressBar.value++;
    pref["glyphTable"]["cid"]["Pr6"]    = makeCidTableSubset(pref["mode"], 7, "Pr6");
    panelObj.panel.ProgressBar.value++;
    pref["glyphTable"]["cid"]["Pr6N"] = makeCidTableSubset(pref["mode"], 11, "Pr6N");
    panelObj.panel.ProgressBar.value++;
    
    panelObj.panel.close();

    //変換工数の計算
    var unicodeTableSize = pref["glyphTable"]["unicode"].length;
    var cidTableSize = pref["glyphTable"]["cid"]["Pr6N"].length;

    var cnt = 0;
    if(pref["fontSet"] == "使用中の全フォント") {
        cnt = unicodeTableSize + app.activeDocument.fonts.length * cidTableSize;
    } else {
        cnt = unicodeTableSize + cidTableSize;
    }
    
    //進捗ダイアログの設定
    var panelObj = new ProgressPanel("字形変換中");
    panelObj.setInstance(cnt, 400);
    panelObj.panel.ProgressLabel.text  = "字形変換中です。完了までしばらくお待ちください……" ;
    panelObj.panel.show();
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
