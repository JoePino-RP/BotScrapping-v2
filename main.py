import process as pr
import Functions_Bot as FB
Issue = FB.ExcelData()

Auxilio = """
Necesitamos ayuda con el Issue 
"""+str(int(Issue[15]))
try:
    pr.run_away()
except:
    FB.EnviarEmail("ERROR","joseph.rojas@exsis.com.co",archivo="",msg=Auxilio)
    FB.cerrar()