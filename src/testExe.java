import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.*;

/**
 * Created by SerP on 07.05.2016.
 */
public class testExe {

    public static void main(String[] args) {
        testCase3();
    }

    public static void testCase3()
    {
        String Serv = "TestNew.TestAppS";
        ActiveXComponent comX = new ActiveXComponent(Serv);
        setEvent(comX);
        boolean test;

        test = getCom(comX);
        test = getCom(comX);

        try {
            if (comX != null) {
                System.out.println("Wating for events ...");
                Thread.sleep(200); // 60 seconds is long enough
                System.out.println("Cleaning up ...");
                //comX.safeRelease();
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            ;//ComThread.Release();
        }
    }

    public static void testCase2()
    {
        String Serv = "TestNew.TestSingl";
        ActiveXComponent comX = new ActiveXComponent(Serv);
        setEvent(comX);
        boolean test;

        test = getCom2(comX);
        test = getCom2(comX);

        try {
            if (comX != null) {
                System.out.println("Wating for events ...");
                Thread.sleep(200); // 60 seconds is long enough
                System.out.println("Cleaning up ...");
                //comX.safeRelease();
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            ;//ComThread.Release();
        }
    }

    public static boolean getCom2(ActiveXComponent comX)
    {
        try {
            System.out.println("\nGet Prop_str str, must be empty");
            Object t1 = comX.getProperty("Prop_str");
            System.out.println("Str Prop_str= " + t1.toString());
            String someString = "aaaam";
            System.out.println("\nSet Prop_str str = \'" + someString+"\'");
            comX.setProperty("Prop_str", new Variant(someString));
            Object t2 = comX.getProperty("Prop_str");
            System.out.println("\nGet Prop_str str: ");
            System.out.println("Str Prop_str= " + t2.toString());
        } catch (Exception e) {
            e.printStackTrace();
        }
        return true;
    }


    // One jacob comp and 2 using
    public static void testCase1() {
        //ActiveXComponent comX = null;
        String Serv = "TestNew.ComO";
        ActiveXComponent comX = new ActiveXComponent(Serv);
        setEvent(comX);
        boolean test;

        test = getCom(comX);
        test = getCom(comX);

        try {
            if (comX != null) {
                System.out.println("Wating for events ...");
                Thread.sleep(20000); // 60 seconds is long enough
                System.out.println("Cleaning up ...");
                //comX.safeRelease();
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            ;//ComThread.Release();
        }
    }

    public static boolean getCom(ActiveXComponent comX)
    {
        try {
            System.out.println("\nGet emsxml str, must be empty");
            Object t1 = comX.getProperty("emsxml");
            System.out.println("Str emsxml= " + t1.toString());
            String someString = "aaaam";
            System.out.println("\nSet emsxml str = \'" + someString+"\'");
            comX.setProperty("emsxml", new Variant(someString));
            Object t2 = comX.getProperty("emsxml");
            System.out.println("\nGet emsxml str: ");
            System.out.println("Str emsxml= " + t2.toString());
        } catch (Exception e) {
            e.printStackTrace();
        }
        return true;
    }

    public static void setEvent(ActiveXComponent comX)
    {
        InvocationProxy proxy = new InvocationProxy() {
            @Override
            public Variant invoke(String methodName, Variant[] targetParameters) {
                System.out.println("*** Event ***: " + methodName + " param: " + targetParameters.toString() );
                return null;
            }
        };
        DispatchEvents de = new DispatchEvents((Dispatch) comX.getObject(), proxy);
    }

    public static void testExe(){

        String path = System.getProperty("java.library.path");
        System.out.println(path);
        System.loadLibrary("jacob-1.18-x86");

        //String Serv = "TestServer.EventTest";
        String Serv = "TestNew.ComO";

        ActiveXComponent xl = new ActiveXComponent(Serv);

        //TestServer.EventTest
        //ActiveXComponent xl = new ActiveXComponent("Project3.TestNew");


        Object xlo = xl.getObject();

        InvocationProxy proxy = new InvocationProxy() {
            @Override
            public Variant invoke(String methodName, Variant[] targetParameters) {
                System.out.println("*** Event ***: " + methodName + " param: " + targetParameters.toString() );
                return null;
            }
        };
        DispatchEvents de = new DispatchEvents((Dispatch) xl.getObject(), proxy);

        try {
            System.out.println("\nGet emsxml str, must be empty");
            Object t1 = xl.getProperty("emsxml");
            System.out.println("Str emsxml= " + t1.toString());
            String someString = "aaaam";
            System.out.println("\nSet emsxml str = \'" + someString+"\'");
            xl.setProperty("emsxml", new Variant(someString));
            Object t2 = xl.getProperty("emsxml");
            System.out.println("\nGet emsxml str: ");
            System.out.println("Str emsxml= " + t2.toString());
        } catch (Exception e) {
            e.printStackTrace();
        }


        try {
            if (xl != null) {
                //System.out.println("Version:"+c.getProperty("Version"));
                /*
                InvocationProxy proxy = new InvocationProxy() {
                    @Override
                    public Variant invoke(String methodName, Variant[] targetParameters) {
                        System.out.println("*** Event ***: " + methodName + " param: " + targetParameters[0].toString() );
                        //return targetParameters[0];
                        return null;
                    }
                };
                DispatchEvents de = new DispatchEvents((Dispatch) xl.getObject(), proxy);
                */
                /*
                c.invoke("OnStatusChanged", new Variant[] {
                        new Variant("aaaa")

                });
                */
                System.out.println("Wating for events ...");
                Thread.sleep(200); // 60 seconds is long enough
                System.out.println("Cleaning up ...");
                xl.safeRelease();
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
