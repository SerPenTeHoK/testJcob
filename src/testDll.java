import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

/**
 * Created by SerP on 07.05.2016.
 */
public class testDll {

    public static void main(String[] args) {
        ActiveXComponent xl = new ActiveXComponent("TestServer");
        Object xlo = xl.getObject();
        try {

            System.out.println("version="+xl.getProperty("Version"));
            System.out.println("version="+ Dispatch.get((Dispatch) xlo, "Version"));
        } catch (Exception e) {
            e.printStackTrace();
        }
            /*
            xl.setProperty("Visible", new Variant(true));
            Object workbooks = xl.getProperty("Workbooks").toDispatch();
            Object workbook = Dispatch.get((Dispatch) workbooks,"Add").toDispatch();
            Object sheet = Dispatch.get((Dispatch) workbook,"ActiveSheet").toDispatch();
            Object a1 = Dispatch.invoke((Dispatch) sheet, "Range", Dispatch.Get,
                    new Object[] {"A1"},
                    new int[1]).toDispatch();
            Object a2 = Dispatch.invoke((Dispatch) sheet, "Range", Dispatch.Get,
                    new Object[] {"A2"},
                    new int[1]).toDispatch();
            Dispatch.put((Dispatch) a1, "Value", "123.456");
            Dispatch.put((Dispatch) a2, "Formula", "=A1*2");
            System.out.println("a1 from excel:"+Dispatch.get((Dispatch) a1, "Value"));
            System.out.println("a2 from excel:"+Dispatch.get((Dispatch) a2, "Value"));
            Variant f = new Variant(false);
            Dispatch.call((Dispatch) workbook, "Close", f);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            xl.invoke("Quit", new Variant[] {});
        }
        */
    }
}
