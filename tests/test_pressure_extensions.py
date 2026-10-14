"""Functional checks for strength, reducers, variable bores, Kv and signed labels."""
import math
import os
import sys
import tempfile
import unittest
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from openpyxl import load_workbook

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/"scripts"))
from build_pressure_workbook import FIRST,LAST,make_workbook,recalculate
from test_pressure_workbook import colebrook_bisection

STRENGTH={"E7":100,"E8":10,"E13":240,"E14":"Numeriske testverdier, ikke materialsertifikat","E17":"Ja","E15":2}
RED={"B7":1,"D7":50,"E7":100}
CASES={
    "strength_basic":{"Rørstyrke":STRENGTH},
    "strength_allowances":{"Rørstyrke":{**STRENGTH,"E9":12.5,"E10":2,"E16":0.8}},
    "strength_safety_four":{"Rørstyrke":{**STRENGTH,"E15":4}},
    "strength_zero_pressure":{"Rørstyrke":{**STRENGTH,"E6":0}},
    "strength_overloaded":{"Rørstyrke":{**STRENGTH,"E6":300}},
    "strength_infeasible":{"Rørstyrke":{**STRENGTH,"E6":2000}},
    "strength_invalid_factor":{"Rørstyrke":{**STRENGTH,"E15":0.5}},
    "strength_missing_documentation":{},
    "reducer_contraction":{"Reduksjoner":RED},
    "reducer_expansion":{"Reduksjoner":{"B7":1,"D7":100,"E7":100}},
    "reducer_roundtrip":{"Reduksjoner":{**RED,"B8":1,"C8":50,"D8":82.5,"E8":100}},
    "reducer_own_k":{"Reduksjoner":{**RED,"H7":0.25}},
    "reducer_crane_published":{"Trykkberegning":{"E11":77.9},"Reduksjoner":{"B7":1,"D7":52.5,"E7":0}},
    "reducer_bad_chain":{"Reduksjoner":{**RED,"B8":1,"C8":70,"D8":40,"E8":100}},
    "reducer_bad_length":{"Reduksjoner":{**RED,"E7":-1}},
    "reducer_laminar_own_k":{"Trykkberegning":{"E24":100,"E25":1},"Bend":{"C7":0,"C8":0},"Reduksjoner":{**RED,"H7":0.3}},
    "reducer_laminar_missing_k":{"Trykkberegning":{"E24":100,"E25":1},"Bend":{"C7":0,"C8":0},"Reduksjoner":RED},
    "sections_mixed":{"Trykkberegning":{"E34":"Rørstrekk-tabell"},"Reduksjoner":RED,"Rørstrekk":{"B7":10,"B8":20,"C8":50}},
    "sections_unknown_diameter":{"Trykkberegning":{"E34":"Rørstrekk-tabell"},"Rørstrekk":{"B8":20,"C8":50}},
    "local_bend_after_reducer":{"Reduksjoner":RED,"Bend":{"L7":50,"C7":2}},
    "kv_verified":{"Ventiler":{"B7":2,"H7":50,"G7":"DN80","E7":"Dokumentert test-Kv ved valgt drift","M7":"Ja"}},
    "kv_manual_override":{"Ventiler":{"B7":2,"C7":5,"H7":50}},
    "kv_unverified":{"Ventiler":{"B7":1,"H7":50,"G7":"DN80","E7":"Testkilde","M7":"Nei"}},
    "signed_chart":{"Trykkberegning":{"E30":0,"E31":100}},
}
BAD_HYDRAULIC={"reducer_bad_chain","reducer_bad_length","reducer_laminar_missing_k","sections_unknown_diameter","kv_unverified"}


class ExtensionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temporary=tempfile.TemporaryDirectory(prefix="pressure-extensions-")
        cls.root=Path(cls.temporary.name)
        source=cls.root/"input";source.mkdir()
        selected=os.environ.get("PRESSURE_EXTENSION_CASES")
        cls.names=set(selected.split(",")) if selected else set(CASES)
        if not cls.names<=set(CASES):raise ValueError("Unknown PRESSURE_EXTENSION_CASES")
        paths=[]
        for name,updates in CASES.items():
            if name not in cls.names:continue
            w=make_workbook()
            for sheet,values in updates.items():
                for cell,value in values.items():w[sheet][cell]=value
            path=source/(name+".xlsx");w.save(path);w.close();paths.append(path)
        workers=min(4,len(paths))
        def calculate(i):recalculate(paths[i::workers],cls.root/"output",cls.root/f"profile-{i}",timeout=1200)
        with ThreadPoolExecutor(max_workers=workers) as pool:list(pool.map(calculate,range(workers)))
        print("Current test outputs:",cls.root/"output",flush=True)

    @classmethod
    def tearDownClass(cls):cls.temporary.cleanup()

    def close_enough(self,actual,expected):
        self.assertIsInstance(actual,(int,float))
        self.assertTrue(math.isclose(actual,expected,rel_tol=3e-6,abs_tol=1e-6),(actual,expected))

    def load(self,name,formulas=False):
        w=load_workbook(self.root/"output"/(name+".xlsx"),data_only=not formulas)
        self.addCleanup(w.close)
        return w

    def test_strength(self):
        if not any(name.startswith("strength_") for name in self.names):self.skipTest("No strength cases selected for this focused run")
        for name in self.names:
            if not name.startswith("strength_"):continue
            with self.subTest(case=name):
                w=self.load(name);s=w["Rørstyrke"]
                if name=="strength_invalid_factor":
                    self.assertTrue(s["E39"].value.startswith("FEIL:"));self.assertIsNone(s["H11"].value);continue
                if name=="strength_missing_documentation":
                    self.assertTrue(s["E39"].value.startswith("MANGLER:"));self.assertIsNone(s["H11"].value);continue
                self.assertEqual(s["E39"].value,"OK")
                p=s["E6"].value/10
                ro=s["E7"].value/2
                effective=s["E8"].value*(1-s["E9"].value/100)-s["E10"].value
                ri=ro-effective
                # Closed-end cylinder equilibrium and principal-stress differences.
                theta=p*(ro**2+ri**2)/(ro**2-ri**2)
                axial=p*ri**2/(ro**2-ri**2)
                vm=math.sqrt(3)*p*ro**2/(ro**2-ri**2)
                allowable=s["E13"].value*s["E16"].value/s["E15"].value
                limit=10*allowable*(1-(ri/ro)**2)/math.sqrt(3)
                for cell,expected in {"E11":effective,"E23":theta,"E24":axial,"E25":-p,"E26":vm,"E27":allowable,"E29":vm/allowable,"E30":limit,"H11":limit}.items():self.close_enough(s[cell].value,expected)
                if p:
                    self.close_enough(s["E28"].value,s["E13"].value*s["E16"].value/vm)
                else:self.assertIn("∞",s["E28"].value)
                if name=="strength_infeasible":self.assertEqual(s["E31"].value,"Ingen elastisk løsning")
                else:
                    required=ro*(1-math.sqrt(1-math.sqrt(3)*p/allowable))
                    self.close_enough(s["E31"].value,required)
                    self.close_enough(s["E32"].value,(required+s["E10"].value)/(1-s["E9"].value/100))
                if name=="strength_overloaded":self.assertIn("IKKE OPPFYLT",s["H22"].value)
                print("PASS strength:",name,flush=True)

    def test_hydraulics(self):
        for name in self.names:
            with self.subTest(case=name):
                w=self.load(name);c=w["Beregning"];m=w["Trykkberegning"];red=w["Reduksjoner"]
                errors=[(s.title,x.coordinate,x.value) for s in w for row in s for x in row if x.data_type=="e"]
                self.assertEqual(errors,[])
                if name in BAD_HYDRAULIC:
                    self.assertTrue(c["B4"].value.startswith("FEIL:"),c["B4"].value)
                    self.assertIsNone(m["H7"].value)
                    print("PASS blocked:",name,flush=True);continue
                self.assertEqual(c["B4"].value,"OK",c["B4"].value)
                rho,flow,mu=m["E21"].value,m["E25"].value/3600,m["E22"].value/1000
                inlet=m["E12"].value/1000
                inlet_v=flow/(math.pi*inlet**2/4)
                outlet=inlet;reducer_loss=0
                for r in range(FIRST,LAST+1):
                    if red[f"B{r}"].value!=1:continue
                    d1,d2=red[f"C{r}"].value/1000,red[f"D{r}"].value/1000
                    length=red[f"E{r}"].value/1000
                    angle=math.pi if length==0 else 2*math.atan(abs(d1-d2)/(2*length))
                    beta=min(d1,d2)/max(d1,d2)
                    if red[f"H{r}"].value is not None:k=red[f"H{r}"].value
                    elif d1>d2:k=(0.8*math.sin(angle/2) if angle<math.pi/4 else 0.5*math.sqrt(math.sin(angle/2)))*(1-beta**2)
                    else:k=(2.6*math.sin(angle/2) if angle<math.pi/4 else 1)*(1-beta**2)**2
                    small_v=flow/(math.pi*min(d1,d2)**2/4)
                    reducer_loss+=k*rho*small_v**2/2
                    self.close_enough(red[f"J{r}"].value,k*rho*small_v**2/2000)
                    outlet=d2
                kinetic=rho/2*((flow/(math.pi*outlet**2/4))**2-inlet_v**2)
                self.close_enough(c["B34"].value,reducer_loss)
                self.close_enough(c["B40"].value,kinetic)
                if name=="reducer_crane_published":self.close_enough(red["G7"].value,0.2729017979998056)
                if name=="reducer_roundtrip":self.close_enough(kinetic,0)
                if name=="reducer_expansion":self.assertLess(c["B40"].value,0)
                if name=="sections_mixed":
                    expected=0
                    for diameter,length in [(inlet,10),(.05,20)]:
                        velocity=flow/(math.pi*diameter**2/4)
                        reynolds=rho*velocity*diameter/mu
                        fd=colebrook_bisection(reynolds,m["E15"].value/1000/diameter)
                        expected+=fd*length/diameter*rho*velocity**2/2
                    self.close_enough(c["B29"].value,expected)
                if name=="local_bend_after_reducer":
                    b=w["Bend"]
                    local=.05;v=flow/(math.pi*local**2/4)
                    fd=colebrook_bisection(rho*v*local/mu,m["E15"].value/1000/local)
                    angle=math.pi/2;ratio=1.5;sine=math.sin(angle/2)
                    k=fd*angle*ratio+(0.1+2.4*fd)*sine+6.6*fd*(math.sqrt(sine)+sine)/ratio**(4*angle/math.pi)
                    self.close_enough(b["G7"].value,k)
                    self.close_enough(b["I7"].value,2*k*rho*v**2/2000)
                if name=="kv_verified":self.close_enough(c["B31"].value,2*100000*(rho/1000)*(36/50)**2)
                if name=="kv_manual_override":self.close_enough(c["B31"].value,10000)
                self.close_enough(c["B37"].value,c["B35"].value+c["B36"].value+kinetic)
                self.close_enough(m["H7"].value,c["B37"].value/100000)
                print("PASS hydraulics:",name,flush=True)

    def test_chart_and_choices(self):
        for name in self.names:
            w=self.load(name,formulas=True)
            chart=w["Trykkberegning"]._charts[0]
            labels=chart.dLbls or chart.series[0].dLbls
            self.assertEqual(labels.dLblPos,"outEnd")
            self.assertTrue(labels.showCatName)
            self.assertFalse(labels.showSerName)
            self.assertTrue(chart.x_axis.delete)
            self.assertIn("Ventilvalg",w.defined_names)
            self.assertTrue(any("A7" in v and v.type=="list" for v in w["Ventiler"].data_validations.dataValidation))
            self.assertLess(max(len(cell.value) for s in w for row in s for cell in row if cell.data_type=="f"),8192)


if __name__=="__main__":unittest.main(verbosity=2)
