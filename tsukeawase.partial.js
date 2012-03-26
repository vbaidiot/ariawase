// XlFileFormat
var xlOpenXMLWorkbookMacroEnabled = 52; //.xlsm
var xlExcel9795 = 43; //.xls 97-2003 format in Excel 2003 or prev
var xlExcel8    = 56; //.xls 97-2003 format in Excel 2007

var acNewDatabaseFormatAccess2000 =  9; //.mdb
var acNewDatabaseFormatAccess2002 = 10; //.mdb
var acNewDatabaseFormatAccess2007 = 12; //.accdb

var vbext_ct_StdModule   = 1;
var vbext_ct_ClassModule = 2;
var vbext_ct_MSForm      = 3;
var vbext_ct_Document    = 100;

var acSysCmdAccessVer = 7;

var acModule = 5;

function createCallMainByExt(office) {
    return {
        xls:   office.excel,
        xlsm:  office.excel,
        mdb:   office.access,
        accdb: office.access
    };
}

var pushOf = createCallMainByExt({
    excel: function(conf) {
        var xlApp, xlBook, xlDefCalc;
        try {
            xlApp = new ActiveXObject("Excel.Application");
            xlApp.DisplayAlerts = false;
            xlApp.EnableEvents  = false;
        try {
            xlBook = xlApp.Workbooks.Open(conf.tbin);
            checkExcelMacroSecurity(xlBook);
            
            WScript.Echo("\n  Target: " + conf.tbin);
            cleanDirectory(conf.tsrc);
            exportVBComponents(xlBook.VBProject.VBComponents, conf.tsrc);
        } finally { if (xlBook != null) xlBook.Close(); }
        } finally { if (xlApp  != null) xlApp.Quit();   }
    },
    
    access: function(conf) {
        var acApp;
        try {
            acApp = new ActiveXObject("Access.Application");
            acApp.Visible = false;
        try {
            acApp.OpenCurrentDatabase(conf.tbin);
            
            WScript.Echo("\n  Target: " + conf.tbin);
            cleanDirectory(conf.tsrc);
            exportVBComponents(acApp.VBE.ActiveVBProject.VBComponents, conf.tsrc);
        } finally { if (acApp.CurrentDB() != null) acApp.CurrentDB().Close(); }
        } finally { if (acApp != null)             acApp.Quit(); }
    }
});

var pullOf = createCallMainByExt({
    excel: function(conf) {
        var xlApp, xlBook;
        try {
            xlApp = new ActiveXObject("Excel.Application");
            xlApp.DisplayAlerts = false;
            xlApp.EnableEvents  = false;
        try {
            xlBook = createOpenWorkbook(xlApp, conf.tbin);
            checkExcelMacroSecurity(xlBook);
            
            // TODO: create -> delete
            cleanVBComponents(xlBook.VBProject.VBComponents);
            
            WScript.Echo("\n  Target: " + conf.tbin);
            importVBComponents(xlBook.VBProject.VBComponents, conf.tsrc,
                createFixExcelObjectTrivial(xlBook));
            xlBook.Save();
        } finally { if (xlBook != null) xlBook.Close(); }
        } finally { if (xlApp  != null) xlApp.Quit();   }
    },
    
    access: function(conf) {
        var acApp;
        try {
            acApp = new ActiveXObject("Access.Application");
            acApp.Visible = false;
        try {
            createOpenCurrentDatabase(acApp, conf.tbin);
            
            // TODO: create -> delete
            cleanVBComponents(acApp.VBE.ActiveVBProject.VBComponents);
            
            WScript.Echo("\n  Target: " + conf.tbin);
            importVBComponents(acApp.VBE.ActiveVBProject.VBComponents, conf.tsrc,
                createRunCmdSaveAllModule(acApp));
        } finally { if (acApp.CurrentDB() != null) acApp.CurrentDB().Close(); }
        } finally { if (acApp != null)             acApp.Quit(); }
    }
});

function dateTimeString(dt) {
    var g = function(y) { return (y < 2000) ? 1900 + y : y; };
    var f = function(n) { return (n < 10) ? "0" + n : n.toString(); };
    return g(dt.getYear())  + f(dt.getMonth() + 1) + f(dt.getDate())
         + f(dt.getHours()) + f(dt.getMinutes())   + f(dt.getSeconds());
}

var compoTypeExt = (function() {
    var dict = new ActiveXObject('Scripting.Dictionary');
    dict.Add(vbext_ct_StdModule,   'bas');
    dict.Add(vbext_ct_ClassModule, 'cls');
    dict.Add(vbext_ct_MSForm,      'frm');  //with 'frx'
    dict.Add(vbext_ct_Document,    'dcls'); //custum extension
    return dict;
})();

function createConfig(jobID) {
    var crr = Fso.GetParentFolderName(WScript.ScriptFullName);
    var conf = {
        bin:    Fso.BuildPath(crr, 'bin'),
      //tbin:   undefined,
        src:    Fso.BuildPath(crr, 'src'),
      //tsrc:   undefined,
        backup: Fso.BuildPath(crr, 'backup')
    };
    conf.configure = function() {
        jobID = jobID.toUpperCase();
        switch (jobID) {
        case 'PUSH':
            if (!Fso.FolderExists(conf.bin)) {
                WScript.Echo("The bin directory (" + conf.bin + ") not exists.");
                return false;
            }
            if (!Fso.FolderExists(conf.src)) Fso.CreateFolder(conf.src);
            break;
        case 'PULL':
            if (!Fso.FolderExists(conf.src)) {
                WScript.Echo("The src directory (" + conf.src + ") not exists.");
                return false;
            }
            if (!Fso.FolderExists(conf.bin)) Fso.CreateFolder(conf.bin);
            break;
        default:
            WScript.Echo("The jobID '" + jobID + "' undefined.");
            return false;
        }
        if (!Fso.FolderExists(conf.backup)) Fso.CreateFolder(conf.backup);
        
        conf.backup = Fso.BuildPath(conf.backup, dateTimeString(new Date()) + jobID);
        Fso.CreateFolder(conf.backup);
        return true;
    };
    return conf;
}

function backupTargetSrc(conf) {
    if (!Fso.FolderExists(conf.tsrc)) return;
    Fso.CopyFolder(conf.tsrc, conf.backup+"\\");
}

function cleanDirectory(dir) {
    if (!Fso.FolderExists(dir)) {
         Fso.CreateFolder(dir);
         return;
    }
    
    for (var fs=new Enumerator(Fso.GetFolder(dir).Files),f=fs.item(); !fs.atEnd(); fs.moveNext(),f=fs.item())
        if (compoTypeExt.Exists(f.Name)) f.Delete;
}

function backupTargetBin(conf) {
    if (!Fso.FileExists(conf.tbin)) return;
    Fso.CopyFile(conf.tbin, conf.backup+"\\");
}

function cleanVBComponents(compos) {
    for (var cs=new Enumerator(compos),c=cs.item(); !cs.atEnd(); cs.moveNext(),c=cs.item()) {
        if (c.Type == vbext_ct_Document)
            c.CodeModule.DeleteLines(1, c.CodeModule.CountOfLines);
        else
            compos.Remove(c);
    }
}

function findCollectionByName(coll, cname) {
    for (var cs=new Enumerator(coll),c=cs.item(); !cs.atEnd(); cs.moveNext(),c=cs.item())
        if (c.Name == cname) return c;
    return null;
}

function importVBComponents(compos, impdir, callbackAfterImport) {
    if (callbackAfterImport == null) callbackAfterImport = function() {};
    
    for (var fs=new Enumerator(Fso.GetFolder(impdir).Files),f=fs.item(); !fs.atEnd(); fs.moveNext(),f=fs.item()) {
        var xname = Fso.GetExtensionName(f.Path);
        if (xname == 'frx') continue;
        
        var bname = Fso.GetBaseName(f.Path);
        if (xname != 'dcls') {
            var x = findCollectionByName(compos, bname);
            if (x != null) compos.Remove(x);
        }
        
        var c = compos.Import(f.Path);
        callbackAfterImport(c, f.Path);
        WScript.Echo("  Improted: " + Fso.GetFileName(f.Path));
        if (xname == 'frm') WScript.Echo("  Improted: " + bname + ".frx");
    }
}

function createRunCmdSaveAllModule(acApp) {
    return function(impCompo) {
        impCompo.Activate();
        acApp.DoCmd.Save(acModule, impCompo.Name);
    }
}

function createFixExcelObjectTrivial(xlBook) {
    return function(impCompo, impPath) {
        var xname = Fso.GetExtensionName(impPath);
        if (xname != 'dcls') return;
        
        var compos = xlBook.VBProject.VBComponents;
        
        var origCompo;
        var cname=impCompo.Name, bname=Fso.GetBaseName(impPath);
        if (cname != bname) {
            origCompo = compos.Item(bname);
        }
        else {
            var sht = xlBook.Worksheets.Add();
            compos  = xlBook.VBProject.VBComponents; // refreash Component collection
            origCompo = compos.Item(sht.CodeName);
            
            var tmpname = "ImportTemp";
            for (var x=findCollectionByName(compos, tmpname); x!=null; tmpname+="1");
            impCompo.Name  = tmpname;
            origCompo.Name = cname;
        }
        
        var imod=impCompo.CodeModule, omod=origCompo.CodeModule;
        omod.DeleteLines(1, omod.CountOfLines);
        omod.AddFromString(imod.Lines(1, imod.CountOfLines));
        compos.Remove(impCompo);
    }
}

function exportVBComponents(compos, expdir) {
    for (var cs=new Enumerator(compos),c=cs.item(); !cs.atEnd(); cs.moveNext(),c=cs.item()) {
        var xname = compoTypeExt.Exists(c.Type) ? compoTypeExt.Item(c.Type) : null;
        if (xname == null) continue;
        
        if (isDirectiveOnly(c.CodeModule)) continue;
        
        var bname = c.Name;
        var txt = Fso.BuildPath(expdir, bname + "." + xname);
        c.Export(txt);
        WScript.Echo("  Exproted: " + Fso.GetFileName(txt));
        if (xname == 'frm') WScript.Echo("  Exproted: " + bname + ".frx");
    }
}

function isDirectiveOnly(codeModule) {
    var ml = codeModule.CountOfLines;
    var dl = codeModule.CountOfDeclarationLines;
    if (ml > dl) return false;
    if (ml < 1)  return true;
    for (var i=0,arr=codeModule.Lines(1, dl).split("\r\n"),len=arr.length; i<len; i++) {
        var s = arr[i].replace(/^\s+|\s+$/g, "");
        if (s != "" && s.charAt(0).toLowerCase() != "o") return false;
    }
    return true;
}

function checkExcelMacroSecurity(xlBook) {
    try {
        xlBook.VBProject;
    }
    catch(e) {
        if (e.number == -2146827284)
            e.description = [e.description, "See also http://support.microsoft.com/kb/813969"].join("\n");
        throw e;
    }
}

function createOpenWorkbook(xlApp, path) {
    var xlFileFormat;
    var vernum = parseInt(xlApp.Version);
    switch (Fso.GetExtensionName(path)) {
    case 'xls':  xlFileFormat = xlExcel9795
                 break;
    case 'xlsm': xlFileFormat = xlOpenXMLWorkbookMacroEnabled;
                 break;
    default:     xlFileFormat = (vernum < 12) ? xlExcel9795 : xlOpenXMLWorkbookMacroEnabled;
                 path        += (vernum < 12) ? '.xls'      : '.xlsm';
                 break;
    }
    
    var xlBook;
    try {
        if (Fso.FileExists(path)) {
            xlBook = xlApp.Workbooks.Open(path);
        }
        else {
            xlBook = xlApp.Workbooks.Add();
            xlBook.SaveAs(path, xlFileFormat);
        }
    }
    catch (ex) {
        if (xlBook != null) xlBook.Close();
        throw ex;
    }
    return xlBook;
}

function createOpenCurrentDatabase(acApp, path) {
    var dbFormat;
    var vernum = parseInt(acApp.SysCmd(acSysCmdAccessVer));
    switch (Fso.GetExtensionName(path)) {
    case 'mdb':   dbFormat = acNewDatabaseFormatAccess2000;
                  break;
    case 'accdb': dbFormat = acNewDatabaseFormatAccess2007;
                  break;
    default:      dbFormat = (vernum < 12) ? acNewDatabaseFormatAccess2002 : acNewDatabaseFormatAccess2007;
                  path    += (vernum < 12) ? '.mdb'                        : '.accdb';
                  break;
    }
    
    if (!Fso.FileExists(path))
        acApp.NewCurrentDatabase(path, dbFormat);
    else
        acApp.OpenCurrentDatabase(path);
    
    return path;
}
