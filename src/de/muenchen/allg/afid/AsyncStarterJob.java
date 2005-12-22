package de.muenchen.allg.afid;

import com.sun.star.beans.NamedValue;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.lib.uno.helper.Factory;
import com.sun.star.lib.uno.helper.WeakBase;
import com.sun.star.task.XAsyncJob;
import com.sun.star.uno.XComponentContext;

public class AsyncStarterJob extends WeakBase implements XServiceInfo,
        XAsyncJob {
    public final XComponentContext ctx;

    protected static final java.lang.String[] SERVICENAMES = { "com.sun.star.task.AsyncJob" };

    public static final java.lang.String IMPLEMENTATIONNAME = "com.sun.star.comp.framework.java.services.AsyncJob";

    public AsyncStarterJob(XComponentContext xCompContext) {
        ctx = xCompContext;
    }

    public synchronized void executeAsync(
            com.sun.star.beans.NamedValue[] lArgs,
            com.sun.star.task.XJobListener xListener)
            throws com.sun.star.lang.IllegalArgumentException {
        // For asynchronous jobs a valid listener reference is guranteed normaly
        // ...
        if (xListener == null)
            throw new com.sun.star.lang.IllegalArgumentException(
                "invalid listener");

        // extract all possible sub list of given argument list
        com.sun.star.beans.NamedValue[] lGenericConfig = null;
        com.sun.star.beans.NamedValue[] lEnvironment = null;

        int c = lArgs.length;
        for (int i = 0; i < c; ++i) {
            if (lArgs[i].Name.equals("Config"))
                lGenericConfig = (com.sun.star.beans.NamedValue[]) com.sun.star.uno.AnyConverter
                    .toArray(lArgs[i].Value);
            else if (lArgs[i].Name.equals("Environment"))
                lEnvironment = (com.sun.star.beans.NamedValue[]) com.sun.star.uno.AnyConverter
                    .toArray(lArgs[i].Value);
        }

        // Analyze the environment info. This sub list is the only guarenteed
        // one!
        if (lEnvironment == null)
            throw new com.sun.star.lang.IllegalArgumentException(
                "no environment");

        java.lang.String sEnvType = null;
        java.lang.String sEventName = null;
        if (sEventName == null) { /* wenigstens einmal sEventName benutzen */
        }
        c = lEnvironment.length;
        for (int i = 0; i < c; ++i) {
            if (lEnvironment[i].Name.equals("EnvType"))
                sEnvType = com.sun.star.uno.AnyConverter
                    .toString(lEnvironment[i].Value);
            else if (lEnvironment[i].Name.equals("EventName"))
                sEventName = com.sun.star.uno.AnyConverter
                    .toString(lEnvironment[i].Value);
        }

        // Further the environment property "EnvType" is required as minimum.
        if ((sEnvType == null)
                || ((!sEnvType.equals("EXECUTOR")) && (!sEnvType
                    .equals("DISPATCH")))) {
            java.lang.String sMessage = "\"" + sEnvType
                    + "\" isn't a valid value for EnvType";
            throw new com.sun.star.lang.IllegalArgumentException(sMessage);
        }

        // Analyze the set of shared config data.
        java.lang.String sAlias = null;
        if (lGenericConfig != null) {
            c = lGenericConfig.length;
            for (int i = 0; i < c; ++i) {
                if (lGenericConfig[i].Name.equals("Alias"))
                    sAlias = com.sun.star.uno.AnyConverter
                        .toString(lGenericConfig[i].Value);
            }
        }
        if (sAlias == null) { /* wenigstens einmal sAlias benutzen */
        }

        // do your job ...
        if (sEventName.equals("onFirstVisibleTask")) {
            onFirstVisibleTask();
        }

        xListener.jobFinished(this, new NamedValue[] {});
    }

    public String[] getSupportedServiceNames() {
        return SERVICENAMES;
    }

    public boolean supportsService(String sService) {
        int len = SERVICENAMES.length;

        for (int i = 0; i < len; i++) {
            if (sService.equals(SERVICENAMES[i]))
                return true;
        }

        return false;
    }

    public String getImplementationName() {
        return (AsyncStarterJob.class.getName());
    }

    public synchronized static com.sun.star.lang.XSingleComponentFactory __getComponentFactory(
            java.lang.String sImplName) {
        com.sun.star.lang.XSingleComponentFactory xFactory = null;
        if (sImplName.equals(AsyncStarterJob.IMPLEMENTATIONNAME))
            xFactory = Factory.createComponentFactory(
                AsyncStarterJob.class, SERVICENAMES);

        return xFactory;
    }

    public synchronized static boolean __writeRegistryServiceInfo(
            com.sun.star.registry.XRegistryKey xRegKey) {
        return Factory.writeRegistryServiceInfo(
            AsyncStarterJob.IMPLEMENTATIONNAME, AsyncStarterJob.SERVICENAMES,
            xRegKey);
    }

    protected void onFirstVisibleTask() {
    };
}