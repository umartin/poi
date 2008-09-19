/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

package org.apache.poi.ss.formula;

import org.apache.poi.hssf.record.formula.AttrPtg;
import org.apache.poi.hssf.record.formula.FuncVarPtg;
import org.apache.poi.hssf.record.formula.Ptg;
import org.apache.poi.hssf.record.formula.function.FunctionMetadataRegistry;
/**
 * Represents a syntactic element from a formula by encapsulating the corresponding <tt>Ptg</tt>
 * token.  Each <tt>ParseNode</tt> may have child <tt>ParseNode</tt>s in the case when the wrapped
 * <tt>Ptg</tt> is non-atomic.
 * 
 * @author Josh Micich
 */
final class ParseNode {

	public static final ParseNode[] EMPTY_ARRAY = { };
	private final Ptg _token;
	private final ParseNode[] _children;
	private boolean _isIf;
	private final int _tokenCount;

	public ParseNode(Ptg token, ParseNode[] children) {
		_token = token;
		_children = children;
		_isIf = isIf(token);
		int tokenCount = 1;
		for (int i = 0; i < children.length; i++) {
			tokenCount += children[i].getTokenCount();
		}
		if (_isIf) {
			// there will be 2 or 3 extra tAttr tokens according to whether the false param is present
			tokenCount += children.length;
		}
		_tokenCount = tokenCount;
	}
	public ParseNode(Ptg token) {
		this(token, EMPTY_ARRAY);
	}
	public ParseNode(Ptg token, ParseNode child0) {
		this(token, new ParseNode[] { child0, });
	}
	public ParseNode(Ptg token, ParseNode child0, ParseNode child1) {
		this(token, new ParseNode[] { child0, child1, });
	}
	private int getTokenCount() {
		return _tokenCount;
	}

	/**
	 * Collects the array of <tt>Ptg</tt> tokens for the specified tree.
	 */
	public static Ptg[] toTokenArray(ParseNode rootNode) {
		TokenCollector temp = new TokenCollector(rootNode.getTokenCount());
		rootNode.collectPtgs(temp);
		return temp.getResult();
	}
	private void collectPtgs(TokenCollector temp) {
		if (isIf(getToken())) {
			collectIfPtgs(temp);
			return;
		}
		for (int i=0; i< getChildren().length; i++) {
			getChildren()[i].collectPtgs(temp);
		}
		temp.add(getToken());
	}
	/**
	 * The IF() function gets marked up with two or three tAttr tokens.
	 * Similar logic will be required for CHOOSE() when it is supported
	 * 
	 * See excelfileformat.pdf sec 3.10.5 "tAttr (19H)
	 */
	private void collectIfPtgs(TokenCollector temp) {

		// condition goes first
		getChildren()[0].collectPtgs(temp);
		
		// placeholder for tAttrIf
		int ifAttrIndex = temp.createPlaceholder();
		
		// true parameter
		getChildren()[1].collectPtgs(temp);
		
		// placeholder for first skip attr
		int skipAfterTrueParamIndex = temp.createPlaceholder();
		int trueParamSize = temp.sumTokenSizes(ifAttrIndex+1, skipAfterTrueParamIndex);

		AttrPtg attrIf = new AttrPtg();
		attrIf.setOptimizedIf(true);
		AttrPtg attrSkipAfterTrue = new AttrPtg();
		attrSkipAfterTrue.setGoto(true);
		
		if (getChildren().length > 2) {
			// false param present
			
			// false parameter
			getChildren()[2].collectPtgs(temp);
			
			int skipAfterFalseParamIndex = temp.createPlaceholder();

			AttrPtg attrSkipAfterFalse = new AttrPtg();
			attrSkipAfterFalse.setGoto(true);

			int falseParamSize =  temp.sumTokenSizes(skipAfterTrueParamIndex+1, skipAfterFalseParamIndex);
			
			attrIf.setData((short)(trueParamSize + 4)); // distance to start of false parameter. +4 for skip after true
			attrSkipAfterTrue.setData((short)(falseParamSize + 4 + 4 - 1)); // 1 less than distance to end of if FuncVar(size=4). +4 for attr skip before 
			attrSkipAfterFalse.setData((short)(4 - 1)); // 1 less than distance to end of if FuncVar(size=4). 

			temp.setPlaceholder(ifAttrIndex, attrIf);
			temp.setPlaceholder(skipAfterTrueParamIndex, attrSkipAfterTrue);
			temp.setPlaceholder(skipAfterFalseParamIndex, attrSkipAfterFalse);
		} else {
			// false parameter not present
			attrIf.setData((short)(trueParamSize + 4)); // distance to start of FuncVar. +4 for skip after true
			attrSkipAfterTrue.setData((short)(4 - 1)); // 1 less than distance to end of if FuncVar(size=4). 
			
			temp.setPlaceholder(ifAttrIndex, attrIf);
			temp.setPlaceholder(skipAfterTrueParamIndex, attrSkipAfterTrue);
		}
		
		temp.add(getToken());
	}

	private static boolean isIf(Ptg token) {
		if (token instanceof FuncVarPtg) {
			FuncVarPtg func = (FuncVarPtg) token;
			if (FunctionMetadataRegistry.FUNCTION_NAME_IF.equals(func.getName())) {
				return true;
			}
		}
		return false;
	}

	public Ptg getToken() {
		return _token;
	}

	public ParseNode[] getChildren() {
		return _children;
	}

	private static final class TokenCollector {

		private final Ptg[] _ptgs;
		private int _offset;

		public TokenCollector(int tokenCount) {
			_ptgs = new Ptg[tokenCount];
			_offset = 0;
		}

		public int sumTokenSizes(int fromIx, int toIx) {
			int result = 0;
			for (int i=fromIx; i<toIx; i++) {
				result += _ptgs[i].getSize();
			}
			return result;
		}

		public int createPlaceholder() {
			return _offset++;
		}

		public void add(Ptg token) {
			if (token == null) {
				throw new IllegalArgumentException("token must not be null");
			}
			_ptgs[_offset] = token;
			_offset++;
		}

		public void setPlaceholder(int index, Ptg token) {
			if (_ptgs[index] != null) {
				throw new IllegalStateException("Invalid placeholder index (" + index + ")");
			}
			_ptgs[index] = token;
		}

		public Ptg[] getResult() {
			return _ptgs;
		}
	}
}