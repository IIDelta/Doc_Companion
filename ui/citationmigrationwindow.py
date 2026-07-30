"""Citation Migration UI. Always migrates a backup copy, never the original."""
from __future__ import annotations
from pathlib import Path
import csv,datetime,traceback
import win32com.client
from PyQt5.QtCore import QObject,QThread,Qt,pyqtSignal
from PyQt5.QtWidgets import QAbstractItemView,QApplication,QFileDialog,QHBoxLayout,QLabel,QMainWindow,QMessageBox,QPushButton,QTableWidget,QTableWidgetItem,QTextEdit,QVBoxLayout,QWidget
from macros.citationmigration import PINK,Plan,PlanItem,Ref,parse_library,plan_sources,zotero_items

class CitationMigrationWindow(QMainWindow):
    def __init__(self,parent=None):
        super().__init__(parent);self.librarypath='';self.library=[];self.plan=None;self.fieldinfo=None;self.setWindowTitle('Citation Migration');self.resize(1100,700);self.ui()
    def ui(self):
        root=QWidget(self);l=QVBoxLayout(root);info=QLabel('Select the shared library export, highlight one citation in Word, then detect and preview a migration. The original document is never changed; a timestamped migration copy is created.');info.setWordWrap(True);l.addWidget(info)
        row=QHBoxLayout();self.liblabel=QLabel('Shared library export: not selected');b=QPushButton('Select Library Export');b.clicked.connect(self.picklibrary);row.addWidget(self.liblabel,1);row.addWidget(b);l.addLayout(row)
        row=QHBoxLayout();self.detectb=QPushButton('Detect Selected Citation');self.detectb.setObjectName('PrimaryButton');self.detectb.clicked.connect(self.detect);self.previewb=QPushButton('Preview Migration');self.previewb.clicked.connect(self.preview);self.migrateb=QPushButton('Create Migration Copy');self.migrateb.setObjectName('SuccessButton');self.migrateb.clicked.connect(self.migrate);self.previewb.setEnabled(False);self.migrateb.setEnabled(False);row.addWidget(self.detectb);row.addWidget(self.previewb);row.addWidget(self.migrateb);row.addStretch();l.addLayout(row)
        self.status=QLabel('Select a library export, then select one citation in Word.');l.addWidget(self.status)
        self.table=QTableWidget(0,5);self.table.setHorizontalHeaderLabels(['Source item','Decision','Replacement','Confidence','Reason']);self.table.setSelectionBehavior(QAbstractItemView.SelectRows);self.table.setEditTriggers(QAbstractItemView.NoEditTriggers);self.table.horizontalHeader().setStretchLastSection(True);l.addWidget(self.table,1)
        self.detail=QTextEdit();self.detail.setReadOnly(True);self.detail.setMaximumHeight(180);l.addWidget(self.detail);self.setCentralWidget(root)
    def picklibrary(self):
        p,_=QFileDialog.getOpenFileName(self,'Select Shared Library Export','','Text Export (*.txt);;All Files (*.*)')
        if p:self.librarypath=p;self.library=parse_library(p);self.liblabel.setText(f'Shared library export: {Path(p).name} ({len(self.library):,} records)')
    def wordselection(self):
        word=win32com.client.GetActiveObject('Word.Application');doc=word.ActiveDocument;sel=word.Selection.Range
        fields=[]
        for f in doc.Fields:
            start=min(f.Code.Start,f.Result.Start);end=max(f.Code.End,f.Result.End)
            if start<sel.End and end>sel.Start:fields.append(f)
        return word,doc,sel,fields
    def detect(self):
        try:
            if not self.library:raise ValueError('Select the shared library export first.')
            word,doc,sel,fields=self.wordselection();text=sel.Text.replace('\r','').replace('\x07','').strip()
            if not text:raise ValueError('Select a citation in Word first.')
            if len(fields)>1:raise ValueError('Selection intersects multiple Word fields. Select one citation field at a time.')
            self.fieldinfo=None
            if fields:
                f=fields[0];code=f.Code.Text or ''
                if 'ADDIN EN.CITE' in code.upper():raise ValueError('The selection contains an existing EndNote field. Use EndNote Edit & Manage Citation(s), not Citation Migration.')
                if 'ZOTERO_ITEM' not in code.upper():raise ValueError('The selected Word field is not a Zotero or EndNote citation field.')
                sources=zotero_items(code);ctype='Zotero field';self.fieldinfo=(min(f.Code.Start,f.Result.Start),max(f.Code.End,f.Result.End),f.Result.Start,f.Result.End)
            else:
                sources=[Ref(raw=text,title=text)];ctype='Plain text'
            items=plan_sources(sources,self.library);self.plan=Plan(ctype,text,items=items);self.fill();self.previewb.setEnabled(True);self.migrateb.setEnabled(True);self.status.setText(f'Detected {ctype}: {len(items)} citation item(s).')
        except Exception as e:QMessageBox.warning(self,'Citation Detection Failed',str(e))
    def fill(self):
        self.table.setRowCount(len(self.plan.items))
        for i,x in enumerate(self.plan.items):
            vals=[x.source.title or x.source.raw,x.decision,x.replacement,f'{x.score:.0%}',x.reason]
            for j,v in enumerate(vals):self.table.setItem(i,j,QTableWidgetItem(v))
        self.table.resizeRowsToContents();self.detail.setPlainText('Planned replacement:\n\n'+self.plan.replacement)
    def preview(self):
        if self.plan:QMessageBox.information(self,'Migration Preview',f'Citation type: {self.plan.citationtype}\n\nCurrent selection:\n{self.plan.originaltext}\n\nReplacement in migration copy:\n{self.plan.replacement}\n\nPink highlighting will be applied to every [REFERENCE REQUIRED: ...] placeholder.')
    def migrate(self):
        try:
            if not self.plan:raise ValueError('Detect a citation first.')
            if not self.fieldinfo and any(x.decision=='Request' for x in self.plan.items):raise ValueError('Unresolved plain-text citations cannot be safely migrated automatically. Resolve the reference first.')
            answer=QMessageBox.question(self,'Create Migration Copy','A backup/migration copy of the Word document will be created and edited. The original file will not be changed. Continue?',QMessageBox.Yes|QMessageBox.No,QMessageBox.No)
            if answer!=QMessageBox.Yes:return
            word,doc,sel,fields=self.wordselection();src=Path(doc.FullName);stamp=datetime.datetime.now().strftime('%Y%m%d%H%M%S');copy=src.with_name(src.stem+'CitationMigration'+stamp+src.suffix);doc.SaveCopyAs(str(copy));copydoc=word.Documents.Open(str(copy))
            if self.fieldinfo:
                fieldstart,fieldend,resultstart,resultend=self.fieldinfo;scope=copydoc.Range(fieldstart,fieldend)
                if scope.Fields.Count<1:raise RuntimeError('Could not find the selected citation field in the migration copy.')
                field=scope.Fields(1);field.Unlink();target=copydoc.Range(resultstart,resultend)
            else:target=copydoc.Range(sel.Start,sel.End)
            replacement=self.plan.replacement
            target.Text=replacement
            # Find unresolved marker text in the final replacement string. This
            # correctly handles braces and semicolons surrounding known EndNote items.
            search_from=0
            for part in self.plan.items:
                if part.decision=='EndNote':
                    continue
                offset=replacement.find(part.replacement,search_from)
                if offset < 0:
                    continue
                rng=copydoc.Range(target.Start+offset,target.Start+offset+len(part.replacement))
                rng.HighlightColorIndex=PINK
                search_from=offset+len(part.replacement)
            copydoc.Save();self.status.setText(f'Migration copy created: {copy.name}');QMessageBox.information(self,'Migration Complete',f'Migration copy created:\n{copy}\n\nRun EndNote → Update Citations and Bibliography in the migration copy. Pink placeholders require follow-up.')
        except Exception as e:QMessageBox.critical(self,'Migration Failed',traceback.format_exc())
