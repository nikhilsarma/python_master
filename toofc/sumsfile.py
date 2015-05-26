Enter file contents hereimport qz
from qz import ui
import dag
from risk.qatools.utils.ui_testdriver import UITestDriver
from risk.qatools.utils.be_tester_tool import BETester
from risk.qatools.utils.app_mappingdriver import MappingApp
from risk.qatools.utils.about_tester_tool import About
from risk.qatools.utils.result_generator import ReportGenerator
import datetime
import xlrd
import sandra
import qztable
import sys
import func
from qz.data.cube import ShowCube
from risk.zinc.api.zn import ZnApi
import qz.lib.io
from risk.zinc.tools.zn_differ import diffQZTables
from risk.zinc.api.zn_dbapi import ZnSqlApi
import risk.zinc.api.zn_dbapi
from risk.zinc.api.zn_filter import uncompress_filter
import qz.data.showtable

class comparejavavssparck(object):

    @dag.cellfn(dag.CanSet)
    def Zn_Domains(self):
        domains = ZnApi.domains('UAT')
        return domains

    @dag.cellfn(dag.CanSet)
    def Zn_Entitled_Domains(self):
        en_domains = ZnApi.entitled_domains('UAT')
        return en_domains

    @dag.cellfn(dag.CanSet)
    def QaSourcePath(self):
        return 'homedirs/home/ZincQA/python;ps'

    @dag.cellfn(dag.CanSet)
    def ProdSourcePath(self):
        return 'ps'

    @dag.cellfn(dag.CanSet)
    def CobDates(self):
        zn = ZnApi('Universal', 'QA')
        return zn.cob_dates()

    @dag.cellfn(dag.CanSet)
    def Cobdate(self):
        cobdate = datetime.date.today() - datetime.timedelta(days=14)
        return cobdate

    def onClickCancel(self, button):
        button.root.close()

    def val(self):
        zn_nocache = ZnApi('universal', 'qa', use_cache=False)
        date1 = self.Cobdate()
        # date1 = '2014-08-14'
        print date1
        snapshot_id = zn_nocache.snapshot_id(date1, 'EOD')
        select_fields = ['SOURCE','SUM("MTM")']
        qzt_qa1 = zn_nocache.query(snapshot_id, select_fields)
        return qzt_qa1

    # @dag.cellfn
    def RunQAQueries(self, *args):
        sub = qz.lib.io.Subprocess(srcdb=self.QaSourcePath())
        # sub = qz.lib.io.Subprocess(srcdb='homedirs/home/navin.k.mishra_clean/python;homedirs/home/ZincQA/python;ps')
        sub.start()
        # f = zn_nocache.query
        # aa = f(snapshot_id, select_fields)
        # sub.start()
        qzt_qa = sub.runFunc(func.partial(self.val))
        sub.stop()
        # qzt_qa = zn_nocache.query(snapshot_id, select_fields)
        print qzt_qa
        return qzt_qa

    # @dag.cellfn
    def RunProdQueries(self, *args):
        sub = qz.lib.io.Subprocess(srcdb=self.ProdSourcePath())
        # sub = qz.lib.io.Subprocess(srcdb='homedirs/home/ZincQA/python;ps')
        sub.start()
        qzt_ps = sub.runFunc(func.partial(self.val))
        # zn_nocache = ZnApi('universal', 'qa', use_cache=False)
        # date = '2014-08-14' #toda
        # snapshot_id = zn_nocache.snapshot_id(date, 'EOD')
        # select_fields = ['SOURCE','SUM("MTM")']
        # qzt_ps = zn_nocache.query(snapshot_id, select_fields)
        # sub.start()
        # f = zn_nocache.query
        # aa = (snapshot_id, select_fields)
        # qzt_ps = sub.runFunc(f, args=aa)
        sub.stop()

        print qzt_ps
        return qzt_ps

    def QztablesCompare(self, *args):
        tbl1=self.RunQAQueries()
        tbl2=self.RunProdQueries()
        rows_only_in_tbl1, rows_only_in_tbl2, field_by_field_diffs, cols_only_in_tbl1, cols_only_in_tbl2, col_type_mismatch_tbl = diffQZTables(tbl1,tbl2,keycols ='SOURCE', ignorecols=None, diff_limit=1000, epsilon=0.0001, strict_schema=False)
        # myTable = qztable.Table(qztable.Schema(['Ticker', 'Price', 'Quantity'], ['string', 'double', 'int32']))
        # myTable.append(['IBM', '204.5', '100'])
        myCube = ShowCube(field_by_field_diffs)
        return [myCube.panel()]

        # return tbl1
        # return ui.CubeQzTableSource(rows_only_in_tbl1)
        # mycube= ShowCube(field_by_field_diffs)
        # ui.VL([ui.Label('Result Pane'), mycube.panel()],size=(ui.Size.STRETCH, ui.Size.DEFAULT))
        # # return ui.VL([ui.Label('Result Pane'), mycube.panel()],size=(ui.Size.STRETCH, ui.Size.DEFAULT))
        # @dag.cellfn
        # def Cube(self):
        # return self.QztablesCompare()

    def onRefresh(self, sender):
        self.LastRefreshedAt.setValue(datetime.datetime.now())

    @dag.cellfn(dag.CanSet)
    def LastRefreshedAt(self):
        return datetime.datetime.now()

    @dag.cellfn
    def ItemsDependency(self):
        self.LastRefreshedAt()

    def comparisontoolpanel(self):
        return ui.VL([
            ui.Spacer( height = 10),
            ui.Label("Welcome to API Regression Tool!!",halign=ui.Align.LEFT, attr=ui.Attr(bold=True,fontSize=9)),
            ui.Separator(),
            ui.Spacer( height = 10),
            ui.HL([
                ui.Label("Enter the QA Source DB Path Here    "),
                ui.AutoCompleteField(
                    value = self.QaSourcePath,
                    allItems = '',
                    showColHeaders = True,
                    addDropDownButton = True,
                    colHeaders=['Existing Paths'],
                    toolTip = 'Please enter the QA Source DB path here.',
                    size=(150, ui.Size.DEFAULT),
                    halign="left"),
                ui.Label("Domains"),
                ui.ComboBox('Universal',self.Zn_Domains,size=(ui.Size.STRETCH, ui.Size.DEFAULT)),
                ui.Label("Entitled Domains"),
                ui.ComboBox('Universal',self.Zn_Entitled_Domains,size=(ui.Size.STRETCH, ui.Size.DEFAULT)),
            ]),
            ui.Spacer(height = 10),
            ui.HL([
                ui.Label("Enter the Prod Source DB Path Here "),
                ui.AutoCompleteField(
                    value = self.ProdSourcePath,
                    allItems = '',
                    showColHeaders = True,
                    addDropDownButton = True,
                    toolTip = 'Please enter the QA Source DB path here.',
                    colHeaders=['Existing Paths'],
                    size=(150, ui.Size.DEFAULT),
                    halign="left"),
                ui.Label("COB Dates"),
                ui.ComboBox(self.Cobdate,self.CobDates,size=(ui.Size.STRETCH, ui.Size.DEFAULT)),
            ]),
            # ui.Spacer( height = 10),
            # ui.HL([
                # ui.Label("Enter the path's start number                       "),
                # ui.AutoCompleteField(
                # toolTip = "Enter the path's start number",
                # value = self.Startpathvalue,
                # allItems = [0,1,2,3,4,5,6,7,8,9,10,11,".",".",".",".",2000],
                # showColHeaders = True,
                # colHeaders=['Paths'],
                # addDropDownButton = True,
                # size=(300,25),
                # halign="left"),
            # ]),
            # ui.Spacer( height = 10),
            # ui.HL([
                # ui.Label("Enter the path's end number                        "),
                # ui.AutoCompleteField(
                # toolTip = "Enter the path's End number",
                # value = self.Endpathvalue,
                # allItems = [0,1,2,3,4,5,6,7,8,9,10,11,".",".",".",".",2000],
                # showColHeaders = True,
                # colHeaders=['Paths'],
                # size=(300,25),
                # addDropDownButton = True,
                # halign="left"),
            # ]),
            ui.Spacer( height = 10),
            ui.HL([
                ui.Spacer( width = 200),
                ui.Button("RUN",onClick=self.onRefresh,glass = True),
                ui.Spacer( width = 5),
                ui.Button("CANCEL",onClick=self.onClickCancel,glass = True),
                # ui.Label('Results for the Comparison are mentioned below:-'),
                # ui.VL(self.QztablesCompareData),
           ]),
            ui.Separator(),
            # ui.VL([ui.Label('Result Pane'), self.QztablesCompare().panel()],size=(ui.Size.STRETCH, ui.Size.DEFAULT)),
            ui.Label('My Cube'),
            ui.Separator(),
            ui.VL(ui.RefreshFunc(self.QztablesCompare,self.ItemsDependency)),
            # self.QztablesCompare(),
        ],halign=ui.Align.CENTER, scroll=ui.Scroll.BOTH)

def main():
    f = ui.Frame(comparejavavssparck().comparisontoolpanel(),size=(ui.Size.STRETCH, ui.Size.DEFAULT),title="API Regression Tool" ,pos = (20,20))
    f.show()
