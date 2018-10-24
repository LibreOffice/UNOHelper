/*
 * Ein einfacher Wrapper für uno-services.
 * Copyright (C) 2005 Christoph Lutz
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is furnished
 * to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE. 
 * 
 * Änderungshistorie der Landeshauptstadt München:
 * 
 * Datum      | Wer | Anderungsgrund
 * -------------------------------------------------------------------
 * 26.04.2005 | BNK | getSimpleName() durch getName() 
 *            |     | ersetzt wg. Java 1.4 Kompatibilität
 * 05.09.2005 | BNK | getImplememtationName() -> getImplementationName() 
 * 09.09.2005 | LUT | +xFilePicker()
 * 09.09.2005 | LUT | BUGFIX: getSortedMethodsIterator holt die Methoden
 *                    jetzt mit getMethods statt mit getDeclaredMethods  
 * 13.09.2005 | LUT | +xFrame()
 * 19.09.2005 | LUT | +xToolkit()
 *                    +xExtendedToolkit()
 * 20.09.2005 | LUT | BUGFIX: introspection wirft NullPointerExceptions in
 *                    getSortedTypesIterator, getSortedServiceIterator,
 *                    getImplementationName
 * 20.09.2005 | LUT | + xWindow()        
 *                    + xWindowPeer()  
 *                    + xComboBox()     
 *                    + xModel()
 *                    + xLayoutManager()
 *                    + xModuleUIConfigurationManagerSupplier()
 *                    + xModuleUIConfigurationManager()
 *                    + create()
 *                    + createWithContext()
 * 30.09.2005 | LUT | + xIndexAccess()
 *                    + xIndexContainer()
 *                    +	xElementAccess()
 *                    +	xDockableWindow()
 *                    +	xListBox()
 *                    +	xVclWindowPeer()
 *                    +	xTextComponent()
 *                    + xModuleManager()
 *                    +	xUIElementSettings()
 *                    +	xConfigManager()
 *                    +	xStringSubstitution()
 *                    +	xSimpleRegistry()
 *                    +	xRegistryKey()
 * 14.10.2005 | LUT | + xBookmarkSupplier()                  
 * 17.10.2005 | LUT | + xNamed()  
 *                    + xDocumentInsertable()                
 *                    + xTextCursor()
 * 18.10.2005 | LUT | + xURLTransformer()
 *                    + xTextRangeCompare()
 * 20.10.2005 | LUT | + xTextFieldSupplier()
 *                    + xEventBroadcaster()
 * 02.11.2005 | LUT | + xComponent()
 * 04.11.2005 | LUT | + xDevice()           
 * 08.11.2005 | LUT | + xDocumentInfoSupplier()           
 *                    + xNameAccess()
 *                    + createWithArguments()
 * 14.11.2005 | LUT | + xModifiable()        
 * 02.01.2006 | BNK | dbg_Services() und dbg_Properties() public
 *                  | dbg_Properties() gibt Werte der Properties aus
 * 03.05.2006 | LUT | + xTextViewCursorSupplier()
 *                    + xTextViewCursor()
 *                    + xDispatchHelper()
 *                    + xDispatchProvider()
 *                    + xController()
 * 18.05.2006 | LUT | + xUpdatable()
 *                    + xTextFramesSupplier()
 * 19.05.2006 | LUT | + xViewSettingsSupplier() 
 * 13.05.2013 | UKT | + Anpassungen an LO 4.0
 * -------------------------------------------------------------------
 */

package de.muenchen.allg.afid;

import java.awt.Button;
import java.awt.Dialog;
import java.awt.FlowLayout;
import java.awt.Frame;
import java.awt.TextArea;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.lang.reflect.Method;
import java.util.Comparator;
import java.util.Iterator;
import java.util.NoSuchElementException;
import java.util.TreeSet;

import com.sun.star.awt.XComboBox;
import com.sun.star.awt.XDevice;
import com.sun.star.awt.XDockableWindow;
import com.sun.star.awt.XExtendedToolkit;
import com.sun.star.awt.XListBox;
import com.sun.star.awt.XTextComponent;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XVclWindowPeer;
import com.sun.star.awt.XWindow;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.beans.Property;
import com.sun.star.beans.XHierarchicalPropertySet;
import com.sun.star.beans.XMultiPropertySet;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertySetInfo;
import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.container.XElementAccess;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XHierarchicalName;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XIndexContainer;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNameContainer;
import com.sun.star.container.XNamed;
import com.sun.star.document.XDocumentInsertable;
import com.sun.star.document.XDocumentPropertiesSupplier;
import com.sun.star.document.XEventBroadcaster;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XConfigManager;
import com.sun.star.frame.XController;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XDispatch;
import com.sun.star.frame.XDispatchHelper;
import com.sun.star.frame.XDispatchProvider;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XLayoutManager;
import com.sun.star.frame.XLayoutManagerEventBroadcaster;
import com.sun.star.frame.XModel;
import com.sun.star.frame.XModuleManager;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.lang.XSingleComponentFactory;
import com.sun.star.lang.XSingleServiceFactory;
import com.sun.star.lang.XTypeProvider;
import com.sun.star.registry.XRegistryKey;
import com.sun.star.registry.XSimpleRegistry;
import com.sun.star.text.XBookmarksSupplier;
import com.sun.star.text.XParagraphCursor;
import com.sun.star.text.XText;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextField;
import com.sun.star.text.XTextFieldsSupplier;
import com.sun.star.text.XTextFramesSupplier;
import com.sun.star.text.XTextRange;
import com.sun.star.text.XTextRangeCompare;
import com.sun.star.text.XTextViewCursor;
import com.sun.star.text.XTextViewCursorSupplier;
import com.sun.star.ui.XModuleUIConfigurationManagerSupplier;
import com.sun.star.ui.XUIConfigurationManager;
import com.sun.star.ui.XUIConfigurationManagerSupplier;
import com.sun.star.ui.XUIConfigurationPersistence;
import com.sun.star.ui.XUIElementFactory;
import com.sun.star.ui.XUIElementSettings;
import com.sun.star.ui.dialogs.XFilePicker;
import com.sun.star.uno.Exception;
import com.sun.star.uno.Type;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.uno.XInterface;
import com.sun.star.util.XChangesBatch;
import com.sun.star.util.XCloseable;
import com.sun.star.util.XModifiable;
import com.sun.star.util.XStringSubstitution;
import com.sun.star.util.XURLTransformer;
import com.sun.star.util.XUpdatable;
import com.sun.star.view.XViewSettingsSupplier;

/**
 * The class UnoService is a wrapper for UnoService-Objects provided by the
 * OpenOffice.org API. The main aim of this wrapper is to make java-programs
 * more readable by avoiding the neccessity of UnoRuntime.queryInterface-calls.
 * The wrapper performs the queryInterface-calls and holds the corresponding
 * interface-instances. Temporary variables for singular interface-types are no
 * longer needed in your own java-programs.
 * 
 * The wrapper provides convenience-methods to certain aspects of uno-services
 * like getter and setter for property-values.
 * 
 * The wrapper also increases transparency when working with uno-objects. It
 * provides methods to inspect uno-services at runtime and print information
 * about supported-services, implemented interfaces, properties and methods. It
 * makes Uno-Programming in java similar to OOo-Basic programming using tools
 * like the famous XRay-tool.
 * 
 * @deprecated 
 */
public class UnoService
{
  /**
   * This field contains the original object of the uno-service.
   */
  private final Object unoObject;

  /**
   * The constructor creates a new wrapperclass for a given uno-service object.
   * 
   * @param unoObject
   *          an arbitrary object returned by the OOo-API
   */
  public UnoService(Object unoObject)
  {
    this.unoObject = unoObject;
  }

  /**
   * This method returns the interface-instance of this UnoService for a
   * specified interface-type. It performs a UnoRuntime.queryInterface call and
   * caches the result interface-type. If the service doesn't implement the
   * required interface, the method returns <code>null</code>. Use the construct
   * <code>if (myUnoService.xAnInterface() != null) ...</code> to ensure the
   * service really implements the interface.
   * 
   * @param ifClass
   *          interface-type to query for.
   * @return The interface instance for this uno-service. If the service doesn't
   *         implement the required interface, the method returns
   *         <code>null</null>.
   */
  public <T> T queryInterface(Class<T> ifClass)
  {
    return UnoRuntime.queryInterface(ifClass, this.unoObject);
  }

  /**********************************************************************************
   * Methods to inspect the uno-service.
   *********************************************************************************/

  /**
   * This method returns the implementationName of this unoService-object.
   * 
   * @return the implementationName of this unoService-object of the String
   *         "noneUnoService" if the object is not a proper UnoService.
   */
  public String getImplementationName()
  {
    if (this.xServiceInfo() != null)
    {
      return this.xServiceInfo().getImplementationName();
    } else
    {
      return "none";
    }
  }

  /**
   * This method provides information about supported services, implemented
   * interfaces, supported properties and methods of this uno-service. The lists
   * of services, implemented interfaces and methods are sorted in alphabetical
   * ascending order.
   * 
   * @return a string containing all the inspection information.
   */
  public String features()
  {
    if (unoObject != null)
    {
      String str = "";
      str += "------------------------------------------\n";
      str += "ImplementationName: " + getImplementationName() + "\n\n";

      str += dbg_Services();
      str += dbg_SupportedInterfaces();
      str += dbg_Properties();
      str += dbg_SupportedInterfacesAndMethods();
      return str;
    } else
    {
      return "null";
    }
  }

  /**
   * This method prints the inspection-information provides by features() to
   * System.out.
   */
  public void printFeatures()
  {
    System.out.println(features());
  }

  /**
   * This method shows the features of a service in a MsgBox.
   */
  public void msgboxFeatures()
  {
    MsgBox.simple("Xray: " + getImplementationName(), features());
  }

  /**
   * This method returns a sorted list of services supported by this
   * uno-service.
   * 
   * @return a string containig the inspection-information.
   */
  public String dbg_Services()
  {
    String str = "Supported UNO-Services: ";
    if (this.xServiceInfo() != null)
    {
      str += "\n";
      Iterator<String> servicesIterator = getSortedServiceIterator();
      while (servicesIterator.hasNext())
      {
        str += "  " + servicesIterator.next() + "\n";
      }
    } else
    {
      str += "none\n";
    }
    return str + "\n";
  }

  /**
   * This method returns a sorted list of properties supported by this
   * uno-service.
   * 
   * @return a string containig the inspection-information.
   */
  public String dbg_Properties()
  {
    String str = "Supported Properties: ";
    if (this.xPropertySet() != null)
    {
      str += "\n";
      Property[] props = this.xPropertySet().getPropertySetInfo()
          .getProperties();
      for (int i = 0; i < props.length; i++)
      {
        String name = props[i].Name;
        str += "  " + name + " - " + props[i].Type.getZClass().getName();
        try
        {
          if ("Bookmark".equals(name))
          {
            UnoService bookmark = this.getPropertyValue(name);
            str += "  (Bookmark '" + bookmark.xNamed().getName() + "')";
          } else
            str += "  (" + this.getPropertyValue(props[i].Name).toString()
                + ")";
        } catch (java.lang.Exception x)
        {
        }
        str += "\n";
      }
    } else
    {
      str += "none\n";
    }
    return str + "\n";
  }

  /**
   * This method returns a sorted list of implemented interfaces.
   * 
   * @return a string containig the inspection-information.
   */
  public String dbg_SupportedInterfaces()
  {
    String str = "Supported Interfaces: ";
    if (this.xTypeProvider() != null)
    {
      str += "\n";
      Iterator<Type> typesIterator = getSortedTypesIterator();
      while (typesIterator.hasNext())
      {
        str += "  " + typesIterator.next().getTypeName() + "\n";
      }
    } else
    {
      str += "none";
    }
    return str + "\n";
  }

  /**
   * This method returns a sorted list of implemented interfaces and their
   * corresponding methods.
   * 
   * @return a string containig the inspection-information.
   */
  private String dbg_SupportedInterfacesAndMethods()
  {
    String str = "Supported Interfaces and Methods: ";
    if (this.xTypeProvider() != null)
    {
      str += "\n";
      Iterator<Type> typesIterator = getSortedTypesIterator();
      while (typesIterator.hasNext())
      {
        Type type = typesIterator.next();
        str += "  " + type.getTypeName() + "\n";
        Iterator<Method> methodsIterator = getSortedMethodsIterator(type);
        while (methodsIterator.hasNext())
        {
          Method m = methodsIterator.next();
          str += "    - " + m.getName() + "(";
          Class<?>[] par = m.getParameterTypes();
          for (int k = 0; k < par.length; k++)
          {
            if (k != 0)
            {
              str += ", ";
            }
            str += par[k].getName();
          }
          str += ") - " + m.getReturnType().getName() + "\n";

        }
      }
    } else
    {
      str += "none";
    }
    return str + "\n";
  }

  /**
   * This method returns an Iterator to a sorted list of supported services.
   */
  private Iterator<String> getSortedServiceIterator()
  {
    if (this.xServiceInfo() != null)
    {
      return getSortedArrayIterator(
          this.xServiceInfo().getSupportedServiceNames(),
          new Comparator<String>()
          {
            @Override
            public int compare(String arg0, String arg1)
            {
              return arg0.compareTo(arg1);
            }

            @Override
            public boolean equals(Object obj)
            {
              return this == obj;
            }
          });
    } else
    {
      return new Iterator<String>()
      {
        @Override
        public void remove()
        {
          throw new IllegalStateException();
        }

        @Override
        public String next()
        {
          throw new NoSuchElementException();
        }

        @Override
        public boolean hasNext()
        {
          return false;
        }
      };
    }
  }

  /**
   * This method returns an Iterator to a sorted list of methods supported by
   * the specified interface-type.
   */
  private Iterator<Method> getSortedMethodsIterator(Type type)
  {
    return getSortedArrayIterator(type.getZClass().getMethods(),
        new Comparator<Method>()
        {
          @Override
          public int compare(Method arg0, Method arg1)
          {
            return arg0.getName().compareTo((arg1).getName());
          }

          @Override
          public boolean equals(Object obj)
          {
            return this == obj;
          }
        });
  }

  /**
   * This method returns an Iterator to a sorted List of implemented
   * interface-types.
   */
  private Iterator<Type> getSortedTypesIterator()
  {
    if (this.xTypeProvider() != null)
    {
      return getSortedArrayIterator(this.xTypeProvider().getTypes(),
          new Comparator<Type>()
          {
            @Override
            public int compare(Type arg0, Type arg1)
            {
              return arg0.getTypeName().compareTo(arg1.getTypeName());
            }

            @Override
            public boolean equals(Object obj)
            {
              return this == obj;
            }
          });
    } else
    {
      return new Iterator<Type>()
      {
        @Override
        public void remove()
        {
          throw new IllegalStateException();
        }

        @Override
        public Type next()
        {
          throw new NoSuchElementException();
        }

        @Override
        public boolean hasNext()
        {
          return false;
        }
      };
    }
  }

  /**
   * This generic method sorts an array comparing its elements using a
   * comperator. The method is needes by the methods getSortedServiceIterator(),
   * getSortedMethodsIterator(Type type) and getSortedTypesIterator(),
   */
  private <T> Iterator<T> getSortedArrayIterator(T[] array, Comparator<T> c)
  {
    TreeSet<T> set = new TreeSet<T>(c);
    for (int i = 0; i < array.length; i++)
    {
      set.add(array[i]);
    }
    return set.iterator();
  }

  /**********************************************************************************
   * convenience-methods
   *********************************************************************************/

  /**
   * This method returns the property of the uno-service and wrappes the result
   * in a new UnoService-object.
   * 
   * @param p
   *          the name of the required property
   * @return the value of the property, wrapped in an UnoService object.
   * @throws Exception
   */
  public UnoService getPropertyValue(String p) throws Exception
  {
    if (this.xPropertySet() != null)
    {
      return new UnoService(this.xPropertySet().getPropertyValue(p));
    } else
    {
      throw new Exception("Service doesn't support interface XPropertySet");
    }
  }

  /**
   * This method sets a property-value of this
   * 
   * @param p
   *          the name of the property to set
   * @param o
   *          the value of the property
   * @throws Exception
   */
  public void setPropertyValue(String p, Object o) throws Exception
  {
    if (this.xPropertySet() != null)
    {
      this.xPropertySet().setPropertyValue(p, o);
    } else
    {
      throw new Exception("Service doesn't support interface XPropertySet");
    }
  }

  /**
   * This method states if this servic-object provides a specified uno-service.
   * 
   * @param s
   *          the fullqualified name of the service.
   * @return true if the object supports the uno-service, otherwise false.
   */
  public boolean supportsService(String s)
  {
    if (this.xServiceInfo() != null)
    {
      return this.xServiceInfo().supportsService(s);
    } else
    {
      return false;
    }
  }

  /**
   * This method returns the original object of this uno-service.
   * 
   * @return original object of this uno-service
   */
  public Object getObject()
  {
    return unoObject;
  }

  /*
   * (non-Javadoc)
   * 
   * @see java.lang.Object#toString()
   */
  @Override
  public String toString()
  {
    return unoObject.toString();
  }

  /**
   * This method creates and returns a new UnoService using the
   * xMultiServiceFactory of the current unoObject.
   * 
   * @param name
   *          the name of the required service
   * @return The new created UnoService. If the service does not exist or the
   *         unoObject doesn't support xMultiServiceFactory, the result is a new
   *         UnoObject(null).
   * @throws Exception
   */
  public UnoService create(String name) throws Exception
  {
    if (this.xMultiServiceFactory() != null)
    {
      return new UnoService(this.xMultiServiceFactory().createInstance(name));
    } else
    {
      return new UnoService(null);
    }
  }

  /**
   * This method creates a new UnoService with arguments using the
   * xMultiServiceFactory of the current unoObject.
   * 
   * @param name
   *          the name of the required service
   * @param args
   *          the used arguments as a array of PropertyValues
   * @return The new created UnoService. If the service does not exist or the
   *         unoObject doesn't support xMultiServiceFactory, the result is a new
   *         UnoObject(null).
   * @throws Exception
   */
  public UnoService create(String name, Object[] args) throws Exception
  {
    if (this.xMultiServiceFactory() != null)
    {
      return new UnoService(
          this.xMultiServiceFactory().createInstanceWithArguments(name, args));
    } else
    {
      return new UnoService(null);
    }
  }

  /**
   * This method creates a new UnoService with a specified context.
   * 
   * @param name
   *          the name of the required service.
   * @param ctx
   *          the context of the new service.
   * @return the new UnoService.
   * @throws Exception
   */
  public static UnoService createWithContext(String name, XComponentContext ctx)
      throws Exception
  {
    return new UnoService(
        ctx.getServiceManager().createInstanceWithContext(name, ctx));
  }

  /**********************************************************************************
   * Interface-Access
   * 
   * Extend this section if you need to access an interface not in this list.
   *********************************************************************************/

  public XTextDocument xTextDocument()
  {
    return queryInterface(XTextDocument.class);
  }

  public XComponentLoader xComponentLoader()
  {
    return queryInterface(XComponentLoader.class);
  }

  public XDesktop xDesktop()
  {
    return queryInterface(XDesktop.class);
  }

  public XText xText()
  {
    return queryInterface(XText.class);
  }

  public XTextRange xTextRange()
  {
    return queryInterface(XTextRange.class);
  }

  public XPropertySet xPropertySet()
  {
    return queryInterface(XPropertySet.class);
  }

  public XCloseable xCloseable()
  {
    return queryInterface(XCloseable.class);
  }

  public XComponentContext xComponentContext()
  {
    return queryInterface(XComponentContext.class);
  }

  public XMultiComponentFactory xMultiComponentFactory()
  {
    return queryInterface(XMultiComponentFactory.class);
  }

  public XMultiServiceFactory xMultiServiceFactory()
  {
    return queryInterface(XMultiServiceFactory.class);
  }

  public XUnoUrlResolver xUnoUrlResolver()
  {
    return queryInterface(XUnoUrlResolver.class);
  }

  public XServiceInfo xServiceInfo()
  {
    return queryInterface(XServiceInfo.class);
  }

  public XPropertySetInfo xPropertySetInfo()
  {
    return queryInterface(XPropertySetInfo.class);
  }

  public XInterface xInterface()
  {
    return queryInterface(XInterface.class);
  }

  public XTypeProvider xTypeProvider()
  {
    return queryInterface(XTypeProvider.class);
  }

  public XEnumerationAccess xEnumerationAccess()
  {
    return queryInterface(XEnumerationAccess.class);
  }

  public XEnumeration xEnumeration()
  {
    return queryInterface(XEnumeration.class);
  }

  public XTextField xTextField()
  {
    return queryInterface(XTextField.class);
  }

  public XTextContent xTextContent()
  {
    return queryInterface(XTextContent.class);
  }

  public XShape xShape()
  {
    return queryInterface(XShape.class);
  }

  public XFilePicker xFilePicker()
  {
    return queryInterface(XFilePicker.class);
  }

  public XFrame xFrame()
  {
    return queryInterface(XFrame.class);
  }

  public XToolkit xToolkit()
  {
    return queryInterface(XToolkit.class);
  }

  public XExtendedToolkit xExtendedToolkit()
  {
    return queryInterface(XExtendedToolkit.class);
  }

  public XWindow xWindow()
  {
    return queryInterface(XWindow.class);
  }

  public XWindowPeer xWindowPeer()
  {
    return queryInterface(XWindowPeer.class);
  }

  public XComboBox xComboBox()
  {
    return queryInterface(XComboBox.class);
  }

  public XModel xModel()
  {
    return queryInterface(XModel.class);
  }

  public XLayoutManager xLayoutManager()
  {
    return queryInterface(XLayoutManager.class);
  }

  public XModuleUIConfigurationManagerSupplier xModuleUIConfigurationManagerSupplier()
  {
    return queryInterface(XModuleUIConfigurationManagerSupplier.class);
  }

  public XUIConfigurationManager xUIConfigurationManager()
  {
    return queryInterface(XUIConfigurationManager.class);
  }

  public XIndexAccess xIndexAccess()
  {
    return queryInterface(XIndexAccess.class);
  }

  public XIndexContainer xIndexContainer()
  {
    return queryInterface(XIndexContainer.class);
  }

  public XElementAccess xElementAccess()
  {
    return queryInterface(XElementAccess.class);
  }

  public XDockableWindow xDockableWindow()
  {
    return queryInterface(XDockableWindow.class);
  }

  public XListBox xListBox()
  {
    return queryInterface(XListBox.class);
  }

  public XVclWindowPeer xVclWindowPeer()
  {
    return queryInterface(XVclWindowPeer.class);
  }

  public XTextComponent xTextComponent()
  {
    return queryInterface(XTextComponent.class);
  }

  public XModuleManager xModuleManager()
  {
    return queryInterface(XModuleManager.class);
  }

  public XUIElementSettings xUIElementSettings()
  {
    return queryInterface(XUIElementSettings.class);
  }

  public XConfigManager xConfigManager()
  {
    return queryInterface(XConfigManager.class);
  }

  public XStringSubstitution xStringSubstitution()
  {
    return queryInterface(XStringSubstitution.class);
  }

  public XSimpleRegistry xSimpleRegistry()
  {
    return queryInterface(XSimpleRegistry.class);
  }

  public XRegistryKey xRegistryKey()
  {
    return queryInterface(XRegistryKey.class);
  }

  public XBookmarksSupplier xBookmarksSupplier()
  {
    return queryInterface(XBookmarksSupplier.class);
  }

  public XNamed xNamed()
  {
    return queryInterface(XNamed.class);
  }

  public XDocumentInsertable xDocumentInsertable()
  {
    return queryInterface(XDocumentInsertable.class);
  }

  public XTextCursor xTextCursor()
  {
    return queryInterface(XTextCursor.class);
  }

  public XURLTransformer xURLTransformer()
  {
    return queryInterface(XURLTransformer.class);
  }

  public XTextRangeCompare xTextRangeCompare()
  {
    return queryInterface(XTextRangeCompare.class);
  }

  public XTextFieldsSupplier xTextFieldsSupplier()
  {
    return queryInterface(XTextFieldsSupplier.class);
  }

  public XEventBroadcaster xEventBroadcaster()
  {
    return queryInterface(XEventBroadcaster.class);
  }

  public XComponent xComponent()
  {
    return queryInterface(XComponent.class);
  }

  public XDevice xDevice()
  {
    return queryInterface(XDevice.class);
  }

  public XDispatch xDispatch()
  {
    return queryInterface(XDispatch.class);
  }

  public XDocumentPropertiesSupplier xDocumentPropertiesSupplier()
  {
    return queryInterface(XDocumentPropertiesSupplier.class);
  }

  public XNameAccess xNameAccess()
  {
    return queryInterface(XNameAccess.class);
  }

  public XModifiable xModifiable()
  {
    return queryInterface(XModifiable.class);
  }

  public XUIElementFactory xUIElementFactory()
  {
    return queryInterface(XUIElementFactory.class);
  }

  public XSingleComponentFactory xSingleComponentFactory()
  {
    return queryInterface(XSingleComponentFactory.class);
  }

  public XUIConfigurationPersistence xUIConfigurationPersistence()
  {
    return queryInterface(XUIConfigurationPersistence.class);
  }

  public XHierarchicalName xHierarchicalName()
  {
    return queryInterface(XHierarchicalName.class);
  }

  public XHierarchicalPropertySet xHierarchicalPropertySet()
  {
    return queryInterface(XHierarchicalPropertySet.class);
  }

  public XMultiPropertySet xMultiPropertySet()
  {
    return queryInterface(XMultiPropertySet.class);
  }

  public XSingleServiceFactory xSingleServiceFactory()
  {
    return queryInterface(XSingleServiceFactory.class);
  }

  public XNameContainer xNameContainer()
  {
    return queryInterface(XNameContainer.class);
  }

  public XChangesBatch xChangesBatch()
  {
    return queryInterface(XChangesBatch.class);
  }

  public XUIConfigurationManagerSupplier xUIConfigurationManagerSupplier()
  {
    return queryInterface(XUIConfigurationManagerSupplier.class);
  }

  public XLayoutManagerEventBroadcaster xLayoutManagerEventBroadcaster()
  {
    return queryInterface(XLayoutManagerEventBroadcaster.class);
  }

  public XParagraphCursor xParagraphCursor()
  {
    return queryInterface(XParagraphCursor.class);
  }

  public XUpdatable xUpdatable()
  {
    return queryInterface(XUpdatable.class);
  }

  public XTextViewCursorSupplier xTextViewCursorSupplier()
  {
    return queryInterface(XTextViewCursorSupplier.class);
  }

  public XTextViewCursor xTextViewCursor()
  {
    return queryInterface(XTextViewCursor.class);
  }

  public XDispatchHelper xDispatchHelper()
  {
    return queryInterface(XDispatchHelper.class);
  }

  public XDispatchProvider xDispatchProvider()
  {
    return queryInterface(XDispatchProvider.class);
  }

  public XController xController()
  {
    return queryInterface(XController.class);
  }

  public XTextFramesSupplier xTextFramesSupplier()
  {
    return queryInterface(XTextFramesSupplier.class);
  }

  public XViewSettingsSupplier xViewSettingsSupplier()
  {
    return queryInterface(XViewSettingsSupplier.class);
  }

  // ... add wrapper-methods for your own interfaces here...

  private static class MsgBox
  {

    private static Dialog d;

    public static void simple(String title, String text)
    {
      Frame window = new Frame();

      d = new Dialog(window, title, true);
      d.setLayout(new FlowLayout(FlowLayout.LEFT));

      Button ok = new Button("OK");
      ok.addActionListener(new ActionListener()
      {
        @Override
        public void actionPerformed(ActionEvent e)
        {
          // Hide dialog
          MsgBox.d.setVisible(false);
        }
      });

      d.add(new TextArea(text, 20, 80));
      d.add(ok);

      d.pack();
      d.setVisible(true);
    }
  }

}
