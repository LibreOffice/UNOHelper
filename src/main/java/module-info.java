/*-
 * #%L
 * UNOHelper
 * %%
 * Copyright (C) 2005 - 2023 The Document Foundation
 * %%
 * Licensed under the EUPL, Version 1.1 or – as soon they will be
 * approved by the European Commission - subsequent versions of the
 * EUPL (the "Licence");
 *
 * You may not use this work except in compliance with the Licence.
 * You may obtain a copy of the Licence at:
 *
 * http://ec.europa.eu/idabc/eupl5
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the Licence is distributed on an "AS IS" basis,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the Licence for the specific language governing permissions and
 * limitations under the Licence.
 * #L%
 */
module org.libreoffice.ext.unohelper {
  exports org.libreoffice.ext.unohelper.common;
  exports org.libreoffice.ext.unohelper.dialog.adapter;
  exports org.libreoffice.ext.unohelper.document.text;
  exports org.libreoffice.ext.unohelper.ui;
  exports org.libreoffice.ext.unohelper.ui.layout;
  exports org.libreoffice.ext.unohelper.util;

  requires transitive org.libreoffice.uno;
  requires transitive org.libreoffice.unoloader;

  requires transitive java.xml;
  requires java.desktop;

  requires org.apache.commons.lang3;
  requires com.google.common;
  requires org.jsoup;
}
