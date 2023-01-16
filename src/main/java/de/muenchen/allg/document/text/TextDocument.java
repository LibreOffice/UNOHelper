/*-
 * #%L
 * UNOHelper
 * %%
 * Copyright (C) 2005 - 2023 Landeshauptstadt München
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
package de.muenchen.allg.document.text;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;

import com.sun.star.text.XTextDocument;

import de.muenchen.allg.afid.UNO;
import de.muenchen.allg.afid.UnoHelperException;
import de.muenchen.allg.afid.UnoProps;
import de.muenchen.allg.util.UnoProperty;

/**
 * Wrapper for {@link XTextDocument}.
 */
public class TextDocument
{
  private XTextDocument doc;

  /**
   * New TextDocument.
   *
   * @param doc
   *          The wrapped UNO-Object.
   */
  public TextDocument(XTextDocument doc)
  {
    this.doc = doc;
  }

  /**
   * Save the document as a PDF file in the temporary system folder.
   *
   * @return The file in the temporary system folder.
   * @throws IOException
   *           Can't create the file.
   * @throws com.sun.star.io.IOException
   *           Can't crate the file.
   * @throws UnoHelperException
   *           Can't create the file.
   */
  public File saveAsTemporaryPDF()
      throws IOException, com.sun.star.io.IOException, UnoHelperException
  {
    File outputFile = Files.createTempFile("WollMux_SLV_", ".pdf").toFile();
    UnoProps props = new UnoProps(UnoProperty.FILTER_NAME, "writer_pdf_Export");
    UNO.XStorable(doc).storeToURL(UNO.convertFilePathToURL(outputFile.getAbsolutePath()), props.getProps());
    return outputFile;
  }
}
