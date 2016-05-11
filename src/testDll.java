import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.*;

/**
 * Created by SerP on 07.05.2016.
 */
public class testDll {

    public static void main(String[] args) {

        String path = System.getProperty("java.library.path");
        System.out.println(path);
        System.loadLibrary("jacob-1.18-x86");

        //ActiveXComponent xl = new ActiveXComponent("Project2.test");
        //TestServer.EventTest
        ActiveXComponent xl = new ActiveXComponent("TestServer.EventTest");
        Object xlo = xl.getObject();
        try {

            Object res = xl.invoke("Method2", " SomeText");

            System.out.printf(res.toString());
            //System.out.println("version="+xl.getProperty("Version"));
            //System.out.println("version="+ Dispatch.get((Dispatch) xlo, "Version"));
        } catch (Exception e) {
            e.printStackTrace();
        }


        try {
            ActiveXComponent c = xl;
            if (c != null) {
                //System.out.println("Version:"+c.getProperty("Version"));
                InvocationProxy proxy = new InvocationProxy() {
                    @Override
                    public Variant invoke(String methodName, Variant[] targetParameters) {
                        System.out.println("*** Event ***: " + methodName + " param: " + targetParameters[0].toString() );
                        //return targetParameters[0];
                        return null;
                    }
                };
                DispatchEvents de = new DispatchEvents((Dispatch) c.getObject(), proxy);
                /*
                c.invoke("OnStatusChanged", new Variant[] {
                        new Variant("aaaa")

                });
                */
                System.out.println("Wating for events ...");
                Thread.sleep(20000); // 60 seconds is long enough
                System.out.println("Cleaning up ...");
                c.safeRelease();
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            ComThread.Release();
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
