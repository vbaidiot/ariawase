// XlFileFormat
var xlExcel9795 = 43; //.xls 97-2003 format in Excel 2003 or prev
var xlExcel8    = 56; //.xls 97-2003 format in Excel 2007
var xlOpenXMLWorkbookMacroEnabled = 52; //.xlsm

// AcNewDatabaseFormat
var acNewDatabaseFormatAccess2000 =  9; //.mdb
var acNewDatabaseFormatAccess2002 = 10; //.mdb
var acNewDatabaseFormatAccess2007 = 12; //.accdb

var vbext_ct_StdModule   = 1;
var vbext_ct_ClassModule = 2;
var vbext_ct_MSForm      = 3;
var vbext_ct_Document    = 100;

// AcSysCmdAction
var acSysCmdAccessVer = 7;

// AcObjectType
var acTable  = 0;
var acQuery  = 1;
var acForm   = 2;
var acReport = 3;
var acMacro  = 4;

var fso = WScript.CreateObject("Scripting.FileSystemObject");

var scriptPath = WScript.ScriptFullName;

var args = (function () {
    var a = new Array(WScript.Arguments.length);
    for (var i = 0; i < a.length; i++) a[i] = WScript.Arguments.item(i);
    return a;
}());

var println = function(str) {
    WScript.Echo(str);
};

var dateTimeString = function(dt) {
    var g = function(y) { return (y < 2000) ? 1900 + y : y; };
    var f = function(n) { return (n < 10) ? "0" + n : n.toString(); };
    return g(dt.getYear())  + f(dt.getMonth() + 1) + f(dt.getDate())
         + f(dt.getHours()) + f(dt.getMinutes())   + f(dt.getSeconds());
};

var Config = function() {
    this.root = fso.GetParentFolderName(scriptPath);
    this.bin  = fso.BuildPath(this.root, 'bin');
    this.src  = fso.BuildPath(this.root, 'src');
};
Config.prototype.getBins = function() { return fso.GetFolder(this.bin).Files; };
Config.prototype.getSrcs = function() { return fso.GetFolder(this.src).SubFolders; };

var conf = new Config();

var Office = function() {};
Office.prototype.compoTypeExt = (function() {
    var fwd = {};
    fwd[vbext_ct_StdModule] =   'bas';
    fwd[vbext_ct_ClassModule] = 'cls';
    fwd[vbext_ct_MSForm] =      'frm'; // with 'frx'
    fwd[vbext_ct_Document] =    'dcm'; // custum extension
    
    var rev = {};
    for (var k in fwd) rev[fwd[k]] = k;
    
    return {
        fwd: fwd,
        rev: rev
    };
}());
Office.prototype.isDirectiveOnly = function(codeModule) {
    var ml = codeModule.CountOfLines;
    var dl = codeModule.CountOfDeclarationLines;
    if (ml > dl) return false;
    if (ml < 1)  return true;
    for (var i=0,arr=codeModule.Lines(1, dl).split("\r\n"),len=arr.length; i<len; i++) {
        var s = arr[i].replace(/^\s+|\s+$/g, "");
        if (s != "" && s.charAt(0).toLowerCase() != "o") return false;
    }
    return true;
};
Office.prototype.findCollectionByName = function(coll, cname) {
    for (var cs=new Enumerator(coll),c=cs.item(); !cs.atEnd(); cs.moveNext(),c=cs.item())
        if (c.Name == cname) return c;
    return null;
};
Office.prototype.cleanupBinary = function(compos) {
    for (var cs=new Enumerator(compos),c=cs.item(); !cs.atEnd(); cs.moveNext(),c=cs.item()) {
        if (c.Type == vbext_ct_Document)
            c.CodeModule.DeleteLines(1, c.CodeModule.CountOfLines);
        else
            compos.Remove(c);
    }
};
Office.prototype.cleanupSource = function(dir) {
    if (!fso.FolderExists(dir)) {
         fso.CreateFolder(dir);
         return;
    }
    
    for (var fs=new Enumerator(fso.GetFolder(dir).Files),f=fs.item(); !fs.atEnd(); fs.moveNext(),f=fs.item())
        if (fso.GetExtensionName(f.Path) in this.compoTypeExt.rev) f.Delete();
};
Office.prototype.createOpenFile = function(app, path) {
    throw new Error("Must override.");
};
Office.prototype.importComponents = function(compos, impdir, importDocuments) {
    for (var fs=new Enumerator(fso.GetFolder(impdir).Files),f=fs.item(); !fs.atEnd(); fs.moveNext(),f=fs.item()) {
        var xname = fso.GetExtensionName(f.Path);
        var bname = fso.GetBaseName(f.Path);
        if (xname == 'frx') continue;
        
        if (xname != 'dcm')
            compos.Import(f.Path);
        else
            importDocuments(f.Path, compos);
        
        println("  Improted: " + fso.GetFileName(f.Path));
        if (xname == 'frm') println("  Improted: " + bname + ".frx");
    }
};
Office.prototype.exportComponents = function(compos, expdir, exportDocuments) {
    for (var cs=new Enumerator(compos),c=cs.item(); !cs.atEnd(); cs.moveNext(),c=cs.item()) {
        var xname = this.compoTypeExt.fwd[c.Type.toString()];
        var bname = c.Name;
        if (this.isDirectiveOnly(c.CodeModule)) continue;
        
        var fname = bname + "." + xname;
        if (xname != 'dcm')
            c.Export(fso.BuildPath(expdir, fname));
        else
            exportDocuments(c, expdir);
        
        println("  Exproted: " + fname);
        if (xname == 'frm') println("  Exproted: " + bname + ".frx");
    }
};
Office.prototype.combine   = function() {};
Office.prototype.decombine = function() {};

var Excel = function() {};
Excel.prototype = new Office();
Excel.prototype.createOpenFile = function(xlApp, path) {
    var xlFileFormat;
    var vernum = parseInt(xlApp.Version);
    switch (fso.GetExtensionName(path)) {
    case 'xls':  xlFileFormat = xlExcel9795;
                 break;
    case 'xlsm': xlFileFormat = xlOpenXMLWorkbookMacroEnabled;
                 break;
    default:     xlFileFormat = (vernum < 12) ? xlExcel9795 : xlOpenXMLWorkbookMacroEnabled;
                 path        += (vernum < 12) ? '.xls'      : '.xlsm';
                 break;
    }
    
    var xlBook;
    try {
        if (fso.FileExists(path)) {
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
};
Excel.prototype.checkExcelMacroSecurity = function(xlBook) {
    try {
        xlBook.VBProject;
    }
    catch (ex) {
        if (ex.number == -2146827284)
            ex.description = [ex.description, "See also http://support.microsoft.com/kb/813969"].join("\n");
        throw ex;
    }
};
Excel.prototype.createImportDocument = function(xlBook) {
    return function(path, compos) {
        var impCompo = compos.Import(path);
        
        var origCompo;
        var cname=impCompo.Name, bname=fso.GetBaseName(path);
        if (cname != bname) {
            origCompo = compos.item(bname);
        }
        else {
            var sht = xlBook.Worksheets.Add();
            compos  = xlBook.VBProject.VBComponents; // refreash Component collection
            origCompo = compos.item(sht.CodeName);
            
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
};
Excel.prototype.loanOfXlBook = function(path, isCreate, callback) {
    var xlApp, xlBook, ret;
    
    try {
        xlApp = new ActiveXObject("Excel.Application");
        xlApp.DisplayAlerts = false;
        xlApp.EnableEvents  = false;
    try {
        xlBook = (isCreate) ? this.createOpenFile(xlApp, path) : xlApp.Workbooks.Open(path);;
        this.checkExcelMacroSecurity(xlBook);
        
        ret = callback(xlBook);
    } finally { if (xlBook != null) xlBook.Close(); }
    } finally { if (xlApp  != null) xlApp.Quit();   }
    
    return ret;
};
Excel.prototype.combine = function(tsrc, tbin) {
    println("\n  Target: " + fso.GetFileName(tbin));
    
    var self = this;
    this.loanOfXlBook(tbin, true, function(xlBook) {
        self.cleanupBinary(xlBook.VBProject.VBComponents);
        self.importComponents(
            xlBook.VBProject.VBComponents, tsrc,
            self.createImportDocument(xlBook));
        xlBook.Save();
    });
};
Excel.prototype.decombine = function(tbin, tsrc) {
    println("\n  Target: " + fso.GetFileName(tbin));
    
    var self = this;
    this.loanOfXlBook(tbin, false, function(xlBook) {
        self.cleanupSource(tsrc);
        self.exportComponents(
            xlBook.VBProject.VBComponents, tsrc,
            function(compo, expdir) { compo.Export(fso.BuildPath(expdir, compo.Name + '.dcm')); });
    });
};
Excel.prototype.clear = function(tbin) {
    println("\n  Target: " + fso.GetFileName(tbin));
    
    var self = this;
    this.loanOfXlBook(tbin, false, function(xlBook) {
        self.cleanupBinary(xlBook.VBProject.VBComponents);
        xlBook.Save();
    });
};

var Access = function() {};
Access.prototype = new Office();
Access.prototype.createOpenFile = function(acApp, path) {
    var dbFormat;
    var vernum = parseInt(acApp.SysCmd(acSysCmdAccessVer));
    switch (fso.GetExtensionName(path)) {
    case 'mdb':   dbFormat = acNewDatabaseFormatAccess2000;
                  break;
    case 'accdb': dbFormat = acNewDatabaseFormatAccess2007;
                  break;
    default:      dbFormat = (vernum < 12) ? acNewDatabaseFormatAccess2002 : acNewDatabaseFormatAccess2007;
                  path    += (vernum < 12) ? '.mdb'                        : '.accdb';
                  break;
    }
    
    if (!fso.FileExists(path))
        acApp.NewCurrentDatabase(path, dbFormat);
    else
        acApp.OpenCurrentDatabase(path);
    
    return path;
};
Access.prototype.createImportDocument = function (acProj) {
    return function(path, compos) {
        var acApp = acProj.Application;
        
        var fname = fso.GetBaseName(path);
        var xname = fso.GetExtensionName(fname);
        var bname = fso.GetBaseName(fname);
        var ty = (xname == 'frm') ? acForm
               : (xname == 'rpt') ? acReport
               : (xname == 'mcr') ? acMacro
               : null;
        acApp.LoadFrom(ty, bname, path);
    };
};
Access.prototype.loanOfAcProj = function(path, isCreate, callback) {
    var acApp, acProj, ret;
    
    try {
        acApp = new ActiveXObject("Access.Application");
        acApp.Visible = false;
    try {
        if (isCreate)
            this.createOpenFile(acApp, path);
        else
            acApp.OpenCurrentDatabase(path);
        
        ret = callback(acApp.CurrentProject);
    } finally { if (acApp.CurrentDB() != null) acApp.CurrentDB().Close(); }
    } finally { if (acApp != null)             acApp.Quit(); }
    
    return ret;
};
Access.prototype.combine = function(tsrc, tbin) {
    println("\n  Target: " + fso.GetFileName(tbin));
    
    var self = this;
    this.loanOfAcProj(tbin, true, function(acProj) {
        var acApp = acProj.Application;
        self.cleanupBinary(acApp.VBE.ActiveVBProject.VBComponents);
        self.importComponents(acProj, tsrc, self.createImportDocument(acProj));
    });
};
Access.prototype.exportDocuments = function(acProj, expdir) {
    var acApp = acProj.Application;
    
    var arr = [ { type: acForm,   ext: '.frm', objects: acProj.AllForms   },
                { type: acReport, ext: '.rpt', objects: acProj.AllReports },
                { type: acMacro,  ext: '.mcr', objects: acProj.AllMacros  } ];
    for (var i = 0; i < arr.length; i++) {
        var ty = arr[i].type, xt = arr[i].ext, objs = arr[i].objects;
        for (var j = 0; j < objs.Count; j++)
            acApp.SaveAsText(ty, objs.item(j).Name, fso.BuildPath(expdir, objs.item(j).Name + xt + ".dcm"));
    }
};
Access.prototype.decombine = function(tbin, tsrc) {
    println("\n  Target: " + fso.GetFileName(tbin));
    
    var self = this;
    this.loanOfAcProj(tbin, true, function(acProj) {
        self.cleanupSource(tsrc);
        self.exportComponents(acProj, tsrc, function(compo, expdir) { /* dummy */ });
        self.exportDocuments(acProj, tsrc);
    });
};
Access.prototype.clear = function() {
    println("\n  Target: " + fso.GetFileName(tbin));
    
    var self = this;
    this.loanOfAcProj(tbin, true, function(acProj) {
        var acApp = acProj.Application;
        self.cleanupBinary(acApp.VBE.ActiveVBProject.VBComponents);
    });
};

var Command = function(helper) {
    this.helper = helper;
};
Command.prototype.helper = null;
Command.prototype.combine = function() {
    this.helper.combineImpl(
        "combine", conf.src, conf.bin,
        function() { return conf.getSrcs(); });
};
Command.prototype.decombine = function() {
    this.helper.combineImpl(
        "decombine", conf.bin, conf.src,
        function() { return conf.getBins(); });
};
Command.prototype.clear = function clear() {
    var prop = "clear", getPaths = function() { return conf.getBins(); };
    var self = this;
    this.helper.iterTarget(getPaths, function(path) {
        self.helper.createOffice(path)[prop](path);
    });
};

var CommandHelper = function() {};
CommandHelper.prototype.createOffice = function(fname) {
    switch (fso.GetExtensionName(fname)) {
    case 'xls':
    case 'xlsm':
        return new Excel();
    case 'mdb':
    case 'accdb':
        return new Access();
    default:
        return new Office();
    }
};
CommandHelper.prototype.isTempFile = function(fname) {
    return fname.substring(0, 2) == '~$';
};
CommandHelper.prototype.iterTarget = function(getPaths, action) {
    for (var fs=new Enumerator(getPaths()),f=fs.item(); !fs.atEnd(); fs.moveNext(),f=fs.item() ) {
        if (this.isTempFile(f.Name)) continue;
        action(f.Path);
    }
};
CommandHelper.prototype.combineImpl = function(prop, fromDir, toDir, getPaths) {
    if (!fso.FolderExists(fromDir)) {
        println("directory '" + fromDir + "' not exists.");
        return;
    }
    
    if (!fso.FolderExists(toDir)) fso.CreateFolder(toDir);
    
    var self = this;
    this.iterTarget(getPaths, function(path) {
        self.createOffice(path)[prop](path, fso.BuildPath(toDir, fso.GetFileName(path)));
    });
};
CommandHelper.prototype.getCommand = function(prop) {
    var cmd = new Command(this);
    return (prop in cmd && cmd[prop] != this)
           ? function() { cmd[prop].apply(cmd, arguments); }
           : undefined;
};

function main(args) {
    var prop = args.shift();
    var cmd  = new CommandHelper().getCommand(prop);
    if (cmd == undefined) {
        println("command '" + prop + "' is undefined.");
        return;
    }
    
    cmd();
}

main(args);

