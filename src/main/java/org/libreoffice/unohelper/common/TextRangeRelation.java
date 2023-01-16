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
package org.libreoffice.unohelper.common;

import com.google.common.collect.ImmutableSet;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.text.XTextRange;
import com.sun.star.text.XTextRangeCompare;

/**
 * Relations between to two {@link XTextRange}s.
 */
public enum TextRangeRelation
{

  /**
   * Textrange can't be compared.
   */
  IMPOSSIBLE,
  /**
   * End of textrange A is before start of B.
   *
   * <pre>
   * AAAAAAAA
   *          BBBBBBBB
   * </pre>
   */
  A_BEFORE_B,
  /**
   * Start of textrange A is before start of B and end of A is between start and end of B.
   *
   * <pre>
   * AAAAAAAA
   *     BBBBBBBB
   * </pre>
   */
  A_OVERLAP_B_START,
  /**
   * Start of textrange A is before start of B and end of A is equal to end ofB.
   *
   * <pre>
   * AAAAAAAA
   *     BBBB
   * </pre>
   */
  A_CONTAINS_B_END,
  /**
   * Start of textrange A is before start of B and end of A is after end of B.
   *
   * <pre>
   * AAAAAAAA
   *   BBBB
   * </pre>
   */
  A_CONTAINS_B,
  /**
   * Start of textrange A is equal to start of B and end of A is after end of B.
   *
   * <pre>
   * AAAAAAAA
   * BBBB
   * </pre>
   */
  A_CONTAINS_B_START,
  /**
   * Start of textrange A is between start and end of B and end of A is after end of B.
   *
   * <pre>
   *     AAAAAAAA
   * BBBBBBBB
   * </pre>
   */
  A_OVERLAB_B_END,
  /**
   * Start of textrange A is after end of B.
   *
   * <pre>
   *          AAAAAAAA
   * BBBBBBBB
   * </pre>
   */
  A_AFTER_B,
  /**
   * Start of textrange A is equal to start of B and end of A is before end of B.
   *
   * <pre>
   * AAAA
   * BBBBBBBB
   * </pre>
   */
  A_IN_B_START,
  /**
   * Start of textrange A is after start of B and end of A is before end of B.
   *
   * <pre>
   *   AAAA
   * BBBBBBBB
   * </pre>
   */
  A_IN_B,
  /**
   * Start of textrange A is after start of B and end of A is equal to end of B.
   *
   * <pre>
   *     AAAA
   * BBBBBBBB
   * </pre>
   */
  A_IN_B_END,
  /**
   * Textrange A and B start and end at the same position.
   *
   * <pre>
   * AAAAAAAA
   * BBBBBBBB
   * </pre>
   */
  A_MATCH_B;

  /**
   * Relations where B doesn't overlap with B.
   */
  public static final ImmutableSet<TextRangeRelation> DISTINCT = ImmutableSet.of(A_BEFORE_B, A_AFTER_B);

  /**
   * Textrange A fully lies in B.
   */
  public static final ImmutableSet<TextRangeRelation> A_CHILD_OF_B = ImmutableSet.of(A_IN_B_START, A_IN_B, A_IN_B_END);

  /**
   * Textrange B fully lies in A.
   */
  public static final ImmutableSet<TextRangeRelation> A_PARENT_OF_B = ImmutableSet.of(A_CONTAINS_B_START, A_CONTAINS_B,
      A_CONTAINS_B_END);

  /**
   * Textrange A starts before or equal to B.
   */
  public static final ImmutableSet<TextRangeRelation> A_LESS_THAN_B = ImmutableSet.of(A_BEFORE_B, A_OVERLAP_B_START,
      A_CONTAINS_B_START, A_CONTAINS_B, A_CONTAINS_B_END);

  /**
   * Textrange A starts after or equal to B.
   */
  public static final ImmutableSet<TextRangeRelation> A_GREATER_THAN_B = ImmutableSet.of(A_OVERLAB_B_END, A_AFTER_B,
      A_IN_B_START, A_IN_B, A_IN_B_END);

  /**
   * Textrange A and B have the same start position.
   */
  public static final ImmutableSet<TextRangeRelation> START_IS_SAME = ImmutableSet.of(A_CONTAINS_B_START, A_IN_B_START,
      A_MATCH_B);

  /**
   * Textrange A and B have the same end position.
   */
  public static final ImmutableSet<TextRangeRelation> END_IS_SAME = ImmutableSet.of(A_CONTAINS_B_END, A_MATCH_B,
      A_IN_B_END);

  /**
   * Compares two {@link XTextRange}s.
   * 
   * @param a
   *          The first text range.
   * @param b
   *          The second text range.
   * @return The relation of the text ranges.
   */
  @SuppressWarnings("squid:S3776")
  public static TextRangeRelation compareTextRanges(XTextRange a, XTextRange b)
  {
    XTextRangeCompare compare = null;
    if (a != null)
    {
      compare = UNO.XTextRangeCompare(a.getText());
    }
    if (compare != null && b != null)
    {
      try
      {
        boolean aCollapsed = compare.compareRegionStarts(a.getStart(), a.getEnd()) == 0;
        boolean bCollapsed = compare.compareRegionStarts(b.getStart(), b.getEnd()) == 0;

        int startStart = compare.compareRegionStarts(a.getStart(), b.getStart());
        int startEnd = compare.compareRegionStarts(a.getStart(), b.getEnd());
        int endStart = compare.compareRegionStarts(a.getEnd(), b.getStart());
        int endEnd = compare.compareRegionStarts(a.getEnd(), b.getEnd());

        if (aCollapsed && bCollapsed)
        {
          if (startStart == 1)
          {
            return A_BEFORE_B;
          }
          if (startStart == 0)
          {
            return A_MATCH_B;
          }
          if (startStart == -1)
          {
            return A_AFTER_B;
          }
        }

        if (aCollapsed)
        {
          if (startStart == 1)
          {
            return A_BEFORE_B;
          }
          if (startStart == 0)
          {
            return A_IN_B_START;
          }
          if (startStart == -1 && startEnd == 1)
          {
            return A_IN_B;
          }
          if (startStart == -1 && startEnd == 0)
          {
            return A_IN_B_END;
          }
          if (startStart == -1 && startEnd == -1)
          {
            return A_AFTER_B;
          }
        }

        if (bCollapsed)
        {
          if (startStart == 1 && endStart == 1)
          {
            return A_BEFORE_B;
          }
          if (startStart == 1 && endStart == 0)
          {
            return A_CONTAINS_B_END;
          }
          if (startStart == 1 && endStart == -1)
          {
            return A_CONTAINS_B;
          }
          if (startStart == 0)
          {
            return A_CONTAINS_B_START;
          }
          if (startStart == -1)
          {
            return A_AFTER_B;
          }
        }

        if (startStart == 1 && startEnd == 1 && endStart == -1 && endEnd == 1)
        {
          return A_OVERLAP_B_START;
        }
        if (startStart == 1 && startEnd == 1 && endStart == -1 && endEnd == 0)
        {
          return A_CONTAINS_B_END;
        }
        if (startStart == 1 && startEnd == 1 && endStart == -1 && endEnd == -1)
        {
          return A_CONTAINS_B;
        }
        if (startStart == 0 && startEnd == 1 && endStart == -1 && endEnd == -1)
        {
          return A_CONTAINS_B_START;
        }
        if (startStart == 0 && startEnd == 1 && endStart == -1 && endEnd == 1)
        {
          return A_IN_B_START;
        }
        if (startStart == 0 && startEnd == 1 && endStart == -1 && endEnd == 0)
        {
          return A_MATCH_B;
        }
        if (startStart == -1 && startEnd == 1 && endStart == -1 && endEnd == -1)
        {
          return A_OVERLAB_B_END;
        }
        if (startStart == -1 && startEnd == 1 && endStart == -1 && endEnd == 0)
        {
          return A_IN_B_END;
        }
        if (startStart == -1 && startEnd == 1 && endStart == -1 && endEnd == 1)
        {
          return A_IN_B;
        }
        if (startStart == 1 && (endStart == 1 || endStart == 0))
        {
          return A_BEFORE_B;
        }
        if ((startEnd == -1 || startEnd == 0) && endEnd == -1)
        {
          return A_AFTER_B;
        }
        throw new UnoHelperRuntimeException(
            String.format("Unexpected compare start_start: %d, start_end: %d, end_start: %d, end_end: %d", startStart,
                startEnd, endStart, endEnd));
      } catch (IllegalArgumentException ex)
      {
        // object can't be compared
      }
    }
    return IMPOSSIBLE;
  }
}
