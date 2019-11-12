/*
 * (c) Copyright Ascensio System SIA 2010-2019
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

"use strict";

(function (undefined) {

    var MIN_JACCARD = 0.34;
    var EXCLUDED_PUNCTUATION = {};
    EXCLUDED_PUNCTUATION[46] = true;
    EXCLUDED_PUNCTUATION[95] = true;
    EXCLUDED_PUNCTUATION[160] = true;
    function CNode(oElement, oParent)
    {
        this.element = oElement;
        this.par = null;
        this.children = [];
        this.depth   = 0;
        this.changes = [];
        this.partner = null;
        this.childidx = null;

        this.hashWords = null;
        if(oParent)
        {
            oParent.addChildNode(this);
        }
    }
    CNode.prototype.getElement = function()
    {
        return this.element;
    };

    CNode.prototype.getNeighbors  = function()
    {
        if(!this.par)
        {
            return [undefined, undefined];
        }
        return [this.par.children[this.childidx - 1],  this.par.children[this.childidx + 1]];
    };

    CNode.prototype.print = function()
    {
        if(this.element.print)
        {
            this.element.print();
        }
    };
    CNode.prototype.getDepth = function()
    {
        return this.depth;
    };
    CNode.prototype.addChange = function(oOperation)
    {
        return this.changes.push(oOperation);
    };
    CNode.prototype.equals = function(oNode)
    {
        if(this.depth === oNode.depth)
        {
            var oParent1, oParent2;
            oParent1 = this.par;
            oParent2 = oNode.par;
            if(oParent1 && !oParent2 || !oParent1 && oParent2)
            {
                return false;
            }
            if(oParent1)
            {
                if(!oParent1.equals(oParent2))
                {
                    return false;
                }
            }
            return this.privateCompareElements(oNode, true);
        }
        return false;
    };

    CNode.prototype.privateCompareElements = function(oNode, bCheckNeighbors)
    {
        var oElement1 = this.element;
        var oElement2 = oNode.element;
        if(oElement1.constructor === oElement2.constructor)
        {
            if(typeof oElement1.Value === "number")
            {
                return oElement1.Value === oElement2.Value;
            }
            if(oElement1 instanceof CTextElement)
            {
                if(bCheckNeighbors && oElement1.isSpaceText() && oElement2.isSpaceText())
                {
                    var aNeighbors1 = this.getNeighbors();
                    var aNeighbors2 = oNode.getNeighbors();
                    if(!aNeighbors1[0] && !aNeighbors2[0] || !aNeighbors1[1] && !aNeighbors2[1])
                    {
                        return true;
                    }
                    if(aNeighbors1[0] && aNeighbors2[0])
                    {
                        if(aNeighbors1[0].privateCompareElements(aNeighbors2[0], false))
                        {
                            return true;
                        }
                    }
                    if(aNeighbors1[1] && aNeighbors2[1])
                    {
                        if(aNeighbors1[1].privateCompareElements(aNeighbors2[1], false))
                        {
                            return true;
                        }
                    }
                    return false;
                }
                else
                {
                    return oElement1.equals(oElement2);
                }
            }
            if(oElement1 instanceof CTable)
            {
                if(oElement1.TableGrid.length !== oElement2.TableGrid.length)
                {
                    return false;
                }
            }
            return true;
        }
        return false;
    };
    CNode.prototype.isLeaf = function()
    {
        return this.children.length === 0;
    };

    CNode.prototype.addChildNode = function(oNode)
    {
        oNode.childidx = this.children.length;
        this.children.push(oNode);
        oNode.depth = this.depth + 1;
        oNode.par = this;
    };

    CNode.prototype.isStructure = function()
    {
        return !this.isLeaf();
    };


    CNode.prototype.forEachDescendant = function(callback, T) {
        this.children.forEach(function(node) {
            node.forEach(callback, T);
        });
    };
    CNode.prototype.forEach = function(callback, T) {
        callback.call(T, this);
        this.children.forEach(function(node) {
            node.forEach(callback, T);
        });
    };

    CNode.prototype.setPartner = function (oNode) {
        this.partner = oNode;
        oNode.partner = this;
        if(this.element instanceof CTextElement)
        {
            return this.element.compareFootnotes(oNode.element);
        }
        return null;
    };

    function CTextElement()
    {
        this.elements = [];
        this.firstRun = null;
        this.lastRun = null;
    }

    CTextElement.prototype.equals = function (other)
    {
        if(this.elements.length !== other.elements.length)
        {
            return false;
        }
        for(var i = 0; i < this.elements.length; ++i)
        {
            if(this.elements[i].constructor !== other.elements[i].constructor)
            {
                return false;
            }
            if(typeof this.elements[i].Value === "number")
            {
                if(this.elements[i].Value !== other.elements[i].Value)
                {
                    return false;
                }
            }
            if(this.elements[i].GraphicObj)
            {
                return false;
            }
        }
        return true;
    };

    CTextElement.prototype.updateHash = function(oHash){
        var aCheckArray = [];
        var bVal = false;
        for(var i = 0; i < this.elements.length; ++i)
        {
            var oElement = this.elements[i];
            if(AscFormat.isRealNumber(oElement.Value))
            {
                aCheckArray.push(oElement.Value);
                bVal = true;
            }
            else
            {
                if(oElement instanceof ParaNewLine)
                {
                    if(oElement.BreakType === AscCommonWord.break_Line)
                    {
                        aCheckArray.push(0x000A);
                    }
                    else
                    {
                        aCheckArray.push(0x21A1);
                    }
                }
                else if(oElement instanceof ParaTab)
                {
                    aCheckArray.push(0x0009);
                }
                else if(oElement instanceof ParaSpace)
                {
                    aCheckArray.push(0x20);
                }
            }
        }
        if(aCheckArray.length > 0)
        {
            oHash.update(aCheckArray);
            if(bVal)
            {
                oHash.countLetters++;
            }
        }
    };

    CTextElement.prototype.print = function ()
    {
        var sResultString = "";
        for(var i = 0; i < this.elements.length; ++i)
        {
            if(this.elements[i] instanceof ParaText)
            {
                sResultString += String.fromCharCode(this.elements[i].Value);
            }
            else if(this.elements[i] instanceof ParaSpace)
            {
                sResultString += " ";
            }
        }
        console.log(sResultString);
    };
    CTextElement.prototype.setFirstRun = function (oRun)
    {
        this.firstRun = oRun;
    };
    CTextElement.prototype.setLastRun = function (oRun)
    {
        this.lastRun = oRun;
    };

    CTextElement.prototype.isSpaceText = function ()
    {
        if(this.elements.length === 1)
        {
            return (this.elements[0].Type === para_Space);
        }
        return false;
    };

    CTextElement.prototype.compareFootnotes = function (oTextElement)
    {
        if(this.elements.length === 1 && oTextElement.elements.length === 1
        && this.elements[0].Type === para_FootnoteReference && oTextElement.elements[0].Type === para_FootnoteReference)
        {
            var oBaseContent = this.elements[0].Footnote;
            var oCompareContent = oTextElement.elements[0].Footnote;
            if(oBaseContent && oCompareContent)
            {
                if(!AscCommon.g_oTableId.Get_ById(oBaseContent.Id))
                {
                    var t = oBaseContent;
                    oBaseContent = oCompareContent;
                    oCompareContent = t;
                }
                return [oBaseContent, oCompareContent];
            }
        }
        return null;
    };

    function CMatching()
    {
        this.Footnotes = {};
    }
    CMatching.prototype.get = function(oNode)
    {
        return oNode.partner;
    };

    CMatching.prototype.put = function(oNode1, oNode2)
    {
        var aFootnotes = oNode1.setPartner(oNode2);
        if(aFootnotes)
        {
            this.Footnotes[aFootnotes[0].Id] = aFootnotes[1];
        }
    };

    function ComparisonOptions()
    {
        this.insertionsAndDeletions = null;
        this.moves = null;
        this.comments = null;
        this.formatting = null;
        this.caseChanges = null;
        this.whiteSpace = null;
        this.tables = null;
        this.headersAndFooters = null;
        this.footNotes  = null;
        this.textBoxes = null;
        this.fields = null;
        this.words = null;
    }
    ComparisonOptions.prototype["getInsertionsAndDeletions"] = ComparisonOptions.prototype.getInsertionsAndDeletions = function(){return this.insertionsAndDeletions !== false;};
    ComparisonOptions.prototype["getMoves"] = ComparisonOptions.prototype.getMoves = function(){return this.moves !== false;};
    ComparisonOptions.prototype["getComments"] = ComparisonOptions.prototype.getComments = function(){return this.comments !== false;};
    ComparisonOptions.prototype["getFormatting"] = ComparisonOptions.prototype.getFormatting = function(){return this.formatting !== false;};
    ComparisonOptions.prototype["getCaseChanges"] = ComparisonOptions.prototype.getCaseChanges = function(){return this.caseChanges !== false;};
    ComparisonOptions.prototype["getWhiteSpace"] = ComparisonOptions.prototype.getWhiteSpace = function(){return this.whiteSpace !== false;};
    ComparisonOptions.prototype["getTables"] = ComparisonOptions.prototype.getTables = function(){return true;/*this.tables !== false;*/};
    ComparisonOptions.prototype["getHeadersAndFooters"] = ComparisonOptions.prototype.getHeadersAndFooters = function(){return this.headersAndFooters !== false;};
    ComparisonOptions.prototype["getFootNotes"] = ComparisonOptions.prototype.getFootNotes = function(){return this.footNotes !== false;};
    ComparisonOptions.prototype["getTextBoxes"] = ComparisonOptions.prototype.getTextBoxes = function(){return this.textBoxes !== false;};
    ComparisonOptions.prototype["getFields"] = ComparisonOptions.prototype.getFields = function(){return this.fields !== false;};
    ComparisonOptions.prototype["getWords"] = ComparisonOptions.prototype.getWords = function(){return true;/* this.words !== false;*/};


    ComparisonOptions.prototype["putInsertionsAndDeletions"] = ComparisonOptions.prototype.putInsertionsAndDeletions = function(v){this.insertionsAndDeletions = v;};
    ComparisonOptions.prototype["putMoves"] = ComparisonOptions.prototype.putMoves = function(v){this.moves = v;};
    ComparisonOptions.prototype["putComments"] = ComparisonOptions.prototype.putComments = function(v){this.comments = v;};
    ComparisonOptions.prototype["putFormatting"] = ComparisonOptions.prototype.putFormatting = function(v){this.formatting = v;};
    ComparisonOptions.prototype["putCaseChanges"] = ComparisonOptions.prototype.putCaseChanges = function(v){this.caseChanges = v;};
    ComparisonOptions.prototype["putWhiteSpace"] = ComparisonOptions.prototype.putWhiteSpace = function(v){this.whiteSpace = v;};
    ComparisonOptions.prototype["putTables"] = ComparisonOptions.prototype.putTables = function(v){this.tables = v;};
    ComparisonOptions.prototype["putHeadersAndFooters"] = ComparisonOptions.prototype.putHeadersAndFooters = function(v){this.headersAndFooters = v;};
    ComparisonOptions.prototype["putFootNotes"] = ComparisonOptions.prototype.putFootNotes = function(v){this.footNotes = v;};
    ComparisonOptions.prototype["putTextBoxes"] = ComparisonOptions.prototype.putTextBoxes = function(v){this.textBoxes = v;};
    ComparisonOptions.prototype["putFields"] = ComparisonOptions.prototype.putFields = function(v){this.fields = v;};
    ComparisonOptions.prototype["putWords"] = ComparisonOptions.prototype.putWords = function(v){this.words = v;};


    function CDocumentComparison(oOriginalDocument, oRevisedDocument, oOptions)
    {
        this.originalDocument = oOriginalDocument;
        this.revisedDocument = oRevisedDocument;
        this.options = oOptions;
        this.api = oOriginalDocument.GetApi();
        this.StylesMap = {};
    }
    CDocumentComparison.prototype.getUserName = function()
    {
        var oCore = this.revisedDocument.Core;
        if(oCore && typeof oCore.lastModifiedBy === "string")
        {
            return  oCore.lastModifiedBy.split(";")[0];
        }
        else
        {
            return "Unknown User";
        }
    };

    CDocumentComparison.prototype.isComparableNodes = function(oNode1, oNode2)
    {
        if(oNode1.element.constructor === oNode2.element.constructor)
        {
            if(oNode1.element instanceof CTable)
            {
                if(oNode1.element.TableGrid.length !== oNode2.element.TableGrid.length)
                {
                    return false;
                }
            }
            return true;
        }
        return false;
    };

    CDocumentComparison.prototype.compareRoots = function(oRoot1, oRoot2)
    {

        var oOrigRoot = this.createNodeFromDocContent(oRoot1, null, null);
        var oRevisedRoot =  this.createNodeFromDocContent(oRoot2, null, null);
        var i, j, key;
        var oEqualMap = {};
        var aBase, aCompare, bOrig = true;
        if(oOrigRoot.children.length <= oRevisedRoot.children.length)
        {
            aBase = oOrigRoot.children;
            aCompare = oRevisedRoot.children;
        }
        else
        {
            bOrig = false;
            aBase = oRevisedRoot.children;
            aCompare = oOrigRoot.children;
        }
        var nStart = (new Date()).getTime();
        var aBase2 = [];
        var aCompare2 = [];
        var oCompareMap = {};
        var bMatchNoEmpty = false;
        for(i = 0; i < aBase.length; ++i)
        {
            var oCurNode =  aBase[i];
            if(oCurNode.hashWords)
            {
                var oCurInfo = {

                    jaccard: 0,
                    map: {},
                    minDiff: 1,
                    intersection: 0
                };
                oEqualMap[oCurNode.element.Id] = oCurInfo;
                for(j = 0; j < aCompare.length; ++j)
                {
                    var oCompareNode = aCompare[j];
                    if(oCompareNode.hashWords && this.isComparableNodes(oCurNode, oCompareNode))
                    {
                        var dJaccard = oCurNode.hashWords.jaccard(oCompareNode.hashWords);
                        if(oCurNode.element instanceof CTable)
                        {
                            dJaccard += MIN_JACCARD;
                        }
                        var dIntersection = dJaccard*(oCurNode.hashWords.count + oCompareNode.hashWords.count)/(1+dJaccard);
                        if(oCurInfo.jaccard <= dJaccard)
                        {
                            if(oCurInfo.jaccard < dJaccard)
                            {
                                oCurInfo.map = {};
                                oCurInfo.minDiff = 1;
                            }
                            oCurInfo.map[oCompareNode.element.Id] = oCompareNode;
                            oCurInfo.jaccard = dJaccard;
                            oCurInfo.intersection = dIntersection;
                            if(dJaccard > 0)
                            {
                                var diffA = oCurNode.hashWords.count - dIntersection/dJaccard;
                                var diffB = oCompareNode.hashWords.count - dIntersection/dJaccard;
                                oCurInfo.minDiff = Math.min(oCurInfo.minDiff, Math.min(Math.abs(diffA), Math.abs(diffB)));
                            }
                        }
                    }
                }
                if(oCurInfo.jaccard > MIN_JACCARD)
                {
                    aBase2.push(oCurNode);
                    for(key in oCurInfo.map)
                    {
                        if(oCurInfo.map.hasOwnProperty(key))
                        {
                            oCompareMap[key] = true;
                            if(oCurNode.hashWords.countLetters > 0 && oCurInfo.map[key].hashWords.countLetters > 0)
                            {
                                bMatchNoEmpty = true;
                            }
                        }
                    }
                }
            }
        }
        for(j = 0; j < aCompare.length; ++j)
        {
            oCompareNode = aCompare[j];
            if(oCompareMap[oCompareNode.element.Id])
            {
                aCompare2.push(oCompareNode);
            }
        }
        var nStart1 = (new Date()).getTime();
        console.log("TIME 1: " + (nStart1 - nStart));
        var oLCS;
        var oThis = this;
        var fLCSCallback = function(x, y) {
            var oOrigNode = oLCS.a[x];
            var oReviseNode = oLCS.b[y];
            var oDiff  = new Diff(oOrigNode, oReviseNode);
            oDiff.equals = function(a, b)
            {
                return a.equals(b);
            };
            var oMatching = new CMatching();
            oDiff.matchTrees(oMatching);
            var oDeltaCollector = new DeltaCollector(oMatching, oOrigNode, oReviseNode);
            oDeltaCollector.forEachChange(function(oOperation){
                oOperation.anchor.base.addChange(oOperation);
            });
            oThis.applyChangesToChildNode(oOrigNode);
            for(var key in oMatching.Footnotes)
            {
                if(oMatching.Footnotes.hasOwnProperty(key))
                {
                    var oBaseFootnotes = AscCommon.g_oTableId.Get_ById(key);
                    var oCompareFootnotes = oMatching.Footnotes[key];
                    if(oBaseFootnotes && oCompareFootnotes)
                    {
                        oThis.compareRoots(oBaseFootnotes, oCompareFootnotes);
                    }
                }
            }
        };
        var fEquals = function(a, b)
        {
            if(oEqualMap[a.element.Id])
            {
                if(oEqualMap[a.element.Id].map[b.element.Id])
                {
                    return true;
                }
            }
            else
            {
                if(oEqualMap[b.element.Id])
                {
                    if(oEqualMap[b.element.Id].map[a.element.Id])
                    {
                        return true;
                    }
                }
            }
            return false;
        };


        if(!bMatchNoEmpty)
        {
            if(bOrig)
            {
                for(i = 0; i < aBase2.length; ++i)
                {
                    if(i !== aBase2[i].childidx)
                    {
                        aBase2.splice(i, aBase2[i].length - i);
                        break;
                    }
                }
                for(i = aCompare2.length - 1; i > -1; i--)
                {
                    if(i !== aCompare2[i].childidx)
                    {
                        aCompare2.splice(0, i + 1);
                        break;
                    }
                }
            }
            else
            {

                for(i = 0; i < aCompare2.length; ++i)
                {
                    if(i !== aCompare2[i].childidx)
                    {
                        aCompare2.splice(i, aCompare2[i].length - i);
                        break;
                    }
                }
                for(i = aBase2.length - 1; i > -1; i--)
                {
                    if(i !== aBase2[i].childidx)
                    {
                        aBase2.splice(0, i + 1);
                        break;
                    }
                }
            }

        }
        if(aBase2.length > 0 && aCompare2.length > 0)
        {
            if(bOrig)
            {
                oLCS = new LCS(aBase2, aCompare2);
            }
            else
            {
                oLCS = new LCS(aCompare2, aBase2);
            }
            oLCS.equals = fEquals;
            oLCS.forEachCommonSymbol(fLCSCallback);
        }


        //compare tables and BlockLvlSdt


        //included paragraphs
        if(bMatchNoEmpty)
        {
            i = 0;
            j = 0;
            oCompareMap = {};

            while(i < aBase.length && j < aCompare.length)
            {
                oCurNode = aBase[i];
                oCompareNode = aCompare[j];
                if(oCurNode.partner && oCompareNode.partner)
                {
                    ++i;
                    ++j;
                }
                else
                {
                    var nStartI = i;
                    var nStartJ = j;
                    var nStartComparIndex = j - 1;
                    var nEndCompareIndex = nStartComparIndex;
                    while(j < aCompare.length && !aCompare[j].partner)
                    {
                        ++j;
                    }
                    nEndCompareIndex = j;
                    if((nEndCompareIndex - nStartComparIndex) > 1)
                    {
                        oCompareMap = {};
                        aBase2.length = 0;
                        aCompare2.length = 0;
                        while (i < aBase.length && !aBase[i].partner)
                        {
                            oCurNode = aBase[i];
                            oCurInfo = oEqualMap[oCurNode.element.Id];
                            if(oCurInfo.minDiff < 0.5)
                            {
                                for(key in oCurInfo.map)
                                {
                                    if(oCurInfo.map.hasOwnProperty(key))
                                    {
                                        oCompareNode = oCurInfo.map[key];
                                        if(oCompareNode.childidx > nStartComparIndex
                                            && oCompareNode.childidx < nEndCompareIndex)
                                        {
                                            oCompareMap[key] = true;
                                            aBase2.push(oCurNode);
                                            aCompare2.push(oCompareNode);
                                        }
                                    }
                                }
                            }
                            ++i;
                        }

                        if(aBase2.length > 0 && aCompare2.length > 0)
                        {
                            if(bOrig)
                            {
                                oLCS = new LCS(aBase2, aCompare2);
                            }
                            else
                            {
                                oLCS = new LCS(aCompare2, aBase2);
                            }
                            oLCS.equals = fEquals;
                            oLCS.forEachCommonSymbol(fLCSCallback);
                        }
                    }
                    i = nStartI;
                    j = nStartJ;
                    while(j < aCompare.length && !aCompare[j].partner)
                    {
                        ++j;
                    }
                    while(i < aBase.length && !aBase[i].partner)
                    {
                        ++i;
                    }
                }
            }
        }

        j = oRevisedRoot.children.length - 1;
        i = oOrigRoot.children.length - 1;
        var aInserContent = [];
        var nRemoveCount = 0;
        for(i = oOrigRoot.children.length - 1; i > -1 ; --i)
        {
            if(!oOrigRoot.children[i].partner)
            {
                this.setElementReviewInfoRecursive(oOrigRoot.children[i].element);
                ++nRemoveCount;
            }
            else
            {
                aInserContent.length = 0;
                for(j = oOrigRoot.children[i].partner.childidx + 1;
                    j < oRevisedRoot.children.length && !oRevisedRoot.children[j].partner; ++j)
                {
                    aInserContent.push(oRevisedRoot.children[j]);
                }
                if(aInserContent.length > 0)
                {
                    this.insertNodesToDocContent(oOrigRoot.element, i + 1 + nRemoveCount, aInserContent);
                }
                nRemoveCount = 0;
            }
        }
        aInserContent.length = 0;
        for(j = 0; j < oRevisedRoot.children.length && !oRevisedRoot.children[j].partner; ++j)
        {
            aInserContent.push(oRevisedRoot.children[j]);
        }
        if(aInserContent.length > 0)
        {
            this.insertNodesToDocContent(oOrigRoot.element, nRemoveCount, aInserContent);
        }

        var nStart2 = (new Date()).getTime();
        console.log("TIME 2: " + (nStart2 - nStart1));

    };

    CDocumentComparison.prototype.compareContentsArrays = function(aContentOrig, aContentRev)
    {
        for(var i = 0; i < aContentOrig.length; ++i)
        {
            if(aContentRev[i])
            {
                this.compareRoots(aContentOrig[i], aContentRev[i]);
            }
            else
            {
                this.setDocContentReviewInfoRecursive(aContentOrig[i]);
            }
        }
    };
    CDocumentComparison.prototype.compare = function()
    {
        this.revisedDocument.UpdateAllSectionsInfo();
        var oThis = this;
        var aImages = AscCommon.pptx_content_loader.End_UseFullUrl();
        var oObjectsForDownload = AscCommon.GetObjectsForImageDownload(aImages);


        console.log("COMPARE 1");
        var oApi = oThis.originalDocument.GetApi(), i;
        var fCallback = function (data) {
            console.log("COMPARE 2");

            var oImageMap = {};
            AscCommon.ResetNewUrls(data, oObjectsForDownload.aUrls, oObjectsForDownload.aBuilderImagesByUrl, oImageMap);


            //TODO: Check locks
            History.Create_NewPoint(1);
            var oldTrackRevisions = oThis.originalDocument.IsTrackRevisions();
            oThis.originalDocument.Start_SilentMode();
            oThis.originalDocument.SetTrackRevisions(false);
            var LogicDocuments = oThis.originalDocument.TrackRevisionsManager.Get_AllChangesLogicDocuments();
            for (var LogicDocId in LogicDocuments)
            {
                var LogicDoc = AscCommon.g_oTableId.Get_ById(LogicDocId);
                if (LogicDoc)
                {
                    LogicDoc.AcceptRevisionChanges(undefined, true);
                }
            }
            oThis.originalDocument.End_SilentMode(false);
            oThis.originalDocument.SetTrackRevisions(true);
            var oldUserId = oApi.DocInfo.get_UserId();
            var oldUserName = oApi.DocInfo.get_UserName();
            var oUserInfo = oApi.DocInfo.UserInfo;
            if(!oApi.DocInfo.UserInfo)
            {
                oApi.DocInfo.UserInfo = new Asc.asc_CUserInfo();
            }
            if(oApi.DocInfo.UserInfo)
            {
                oApi.DocInfo.UserInfo.put_Id("");
                oApi.DocInfo.UserInfo.put_FullName(oThis.getUserName());
            }


            oThis.revisedDocument.FieldsManager = oThis.originalDocument.FieldsManager;
            var NewNumbering = oThis.revisedDocument.Numbering.CopyAllNums(oThis.originalDocument.Numbering);
            oThis.revisedDocument.CopyNumberingMap = NewNumbering.NumMap;
            oThis.originalDocument.Numbering.AppendAbstractNums(NewNumbering.AbstractNum);
            oThis.originalDocument.Numbering.AppendNums(NewNumbering.Num);


            oThis.compareRoots(oThis.originalDocument, oThis.revisedDocument);
            var oSectInfoOrig = oThis.originalDocument.SectionsInfo;
            var oSectInfoRevised = oThis.revisedDocument.SectionsInfo;
            if(oSectInfoOrig && oSectInfoRevised)
            {
                var aFooterDefault = [];
                var aFooterEven  = [];
                var aFooterFirst =  [];
                var aHeaderDefault = [];
                var aHeaderEven = [];
                var aHeaderFirst =  [];

                var aFooterDefaultRev =  [];
                var aFooterEvenRev  = [];
                var aFooterFirstRev =  [];
                var aHeaderDefaultRev = [];
                var aHeaderEvenRev = [];
                var aHeaderFirstRev =  [];
                for(i = 0; i < oSectInfoOrig.Elements.length; i++ )
                {
                    var oSectPrOrig = oSectInfoOrig.Elements[i].SectPr;
                    if(oSectPrOrig.FooterDefault)
                    {
                        aFooterDefault.push(oSectPrOrig.FooterDefault.Content);
                    }
                    if(oSectPrOrig.FooterEven)
                    {
                        aFooterEven.push(oSectPrOrig.FooterEven.Content);
                    }
                    if(oSectPrOrig.FooterFirst)
                    {
                        aFooterFirst.push(oSectPrOrig.FooterFirst.Content);
                    }
                    if(oSectPrOrig.HeaderDefault)
                    {
                        aHeaderDefault.push(oSectPrOrig.HeaderDefault.Content);
                    }
                    if(oSectPrOrig.HeaderEven)
                    {
                        aHeaderEven.push(oSectPrOrig.HeaderEven.Content);
                    }
                    if(oSectPrOrig.HeaderFirst)
                    {
                        aHeaderFirst.push(oSectPrOrig.HeaderFirst.Content);
                    }
                }
                for(i = 0; i < oSectInfoRevised.Elements.length; i++ )
                {
                    var oSectPrRev = oSectInfoRevised.Elements[i].SectPr;
                    if(oSectPrRev.FooterDefault)
                    {
                        aFooterDefaultRev.push(oSectPrRev.FooterDefault.Content);
                    }
                    if(oSectPrRev.FooterEven)
                    {
                        aFooterEvenRev.push(oSectPrRev.FooterEven.Content);
                    }
                    if(oSectPrRev.FooterFirst)
                    {
                        aFooterFirstRev.push(oSectPrRev.FooterFirst.Content);
                    }
                    if(oSectPrRev.HeaderDefault)
                    {
                        aHeaderDefaultRev.push(oSectPrRev.HeaderDefault.Content);
                    }
                    if(oSectPrRev.HeaderEven)
                    {
                        aHeaderEvenRev.push(oSectPrRev.HeaderEven.Content);
                    }
                    if(oSectPrRev.HeaderFirst)
                    {
                        aHeaderFirstRev.push(oSectPrRev.HeaderFirst.Content);
                    }
                }
                oThis.compareContentsArrays(aFooterDefault, aFooterDefaultRev);
                oThis.compareContentsArrays(aFooterEven, aFooterEvenRev);
                oThis.compareContentsArrays(aFooterFirst, aFooterFirstRev);
                oThis.compareContentsArrays(aHeaderDefault, aHeaderDefaultRev);
                oThis.compareContentsArrays(aHeaderEven, aHeaderEvenRev);
                oThis.compareContentsArrays(aHeaderFirst, aHeaderFirstRev);
            }
            if(oApi.DocInfo.UserInfo)
            {
                oApi.DocInfo.UserInfo.put_Id(oldUserId);
                oApi.DocInfo.UserInfo.put_FullName(oldUserName);
            }

            oApi.DocInfo.UserInfo = oUserInfo;
            oThis.originalDocument.SetTrackRevisions(oldTrackRevisions);
            oThis.originalDocument.Document_UpdateSelectionState();
            oThis.originalDocument.Document_UpdateInterfaceState();
            oThis.originalDocument.Document_UpdateUndoRedoState();
            var oFonts = oThis.originalDocument.Document_Get_AllFontNames();
            var aFonts = [];
            for (i in oFonts)
            {
                if(oFonts.hasOwnProperty(i))
                {
                    aFonts[aFonts.length] = new AscFonts.CFont(i, 0, "", 0, null);
                }
            }
            oApi.pre_Paste(aFonts, oImageMap, function()
            {
                oThis.originalDocument.Recalculate();
                oApi.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.SlowOperation);
            });

        };
        AscCommon.sendImgUrls(this.originalDocument.Api, oObjectsForDownload.aUrls, fCallback, null, true);
        return null;
    };
    CDocumentComparison.prototype.setTableReviewInfoRecursive = function(oTable)
    {
        for(var i = 0; i < oTable.Content.length; ++i)
        {
            this.setTableRowReviewInfoRecursive(oTable.Content[i]);
        }
    };

    CDocumentComparison.prototype.setTableRowReviewInfoRecursive = function(oRow)
    {
        var j;
        this.updateReviewInfo(oRow, reviewtype_Remove);
        for(j = 0; j < oRow.Content.length; ++j)
        {
            this.setDocContentReviewInfoRecursive(oRow.Content[j].Content);
        }
    };

    CDocumentComparison.prototype.setDocContentReviewInfoRecursive = function(oContent)
    {
        var i, oElement;
        for(i = 0; i < oContent.Content.length; ++i)
        {
            oElement = oContent.Content[i];
            this.setElementReviewInfoRecursive(oElement);
        }
    };
    CDocumentComparison.prototype.setParagraphReviewInfoRecursive = function(oParagraph)
    {
        oParagraph.SelectAll(1);
        var oOldSectPr = oParagraph.SectPr;
        oParagraph.SectPr = undefined;
        oParagraph.Remove(-1);
        oParagraph.SectPr = oOldSectPr;

        if(oParagraph.LogicDocument)
        {
            oParagraph.LogicDocument.ForceCopySectPr = false;
        }
        if(oParagraph.Content[oParagraph.Content.length - 1])
        {
            var oLastRun = oParagraph.Content[oParagraph.Content.length - 1];
            if(oLastRun.ReviewType === reviewtype_Common)
            {
                this.updateReviewInfo(oLastRun, reviewtype_Remove, true);
            }
        }
        var oThis = this;
        oParagraph.CheckRunContent(function (oRun) {
            if(Array.isArray(oRun.Content))
            {
                for(var i = 0; i < oRun.Content.length; ++i)
                {
                    if(oRun.Content[i].Type === para_Drawing)
                    {
                        var aContents = oRun.Content[i].GetAllDocContents();
                        for(var j = 0; j < aContents.length; ++j)
                        {
                            oThis.setDocContentReviewInfoRecursive(aContents[j]);
                        }
                    }
                    else if(oRun.Content[i].Type === para_FootnoteReference)
                    {
                        oThis.setDocContentReviewInfoRecursive(oRun.Content[i].Footnote);
                    }
                }
            }
        });
    };

    CDocumentComparison.prototype.applyChangesToParagraph = function(oNode)
    {
        var oElement = oNode.element, oChange, i, j, k, t, oChildElement,
            oChildNode, oLastText, oFirstText, oCurRun, oNewRun, oFirstRun;
        var oParentParagraph, aContentToInsert;
        oNode.changes.sort(function(c1, c2){return c2.anchor.index - c1.anchor.index});
        for(i = 0; i < oNode.changes.length; ++i)
        {
            oChange = oNode.changes[i];
            oLastText = null;
            oFirstText = null;
            aContentToInsert = [];
            if(oChange.insert.length > 0)
            {
                oFirstText = oChange.insert[0].element;
                oLastText = oChange.insert[oChange.insert.length - 1].element;

                var oLastRemoveText = null;
                var oFirstRemoveText = null;
                if(oChange.remove.length > 0)
                {
                    if(oChange.remove[oChange.remove.length - 1].element instanceof CTextElement)
                    {
                        oLastRemoveText = oChange.remove[oChange.remove.length - 1].element;
                    }
                    if(oChange.remove[0].element instanceof CTextElement)
                    {
                        oFirstRemoveText = oChange.remove[0].element;
                    }
                }
                oCurRun = oLastText.lastRun ? oLastText.lastRun : oLastText;
                oFirstRun = oFirstText.firstRun ? oFirstText.firstRun : oFirstText;
                oParentParagraph =  (oNode.partner && oNode.partner.element) || oCurRun.Paragraph;
                var oParentParagraph2 =  oCurRun.Paragraph;
                for(k = oParentParagraph.Content.length - 1; k > -1; --k)
                {
                    if(oCurRun === oParentParagraph.Content[k])
                    {
                        if(oCurRun instanceof ParaRun)
                        {
                            for(t = oCurRun.Content.length - 1; t > -1; --t)
                            {
                                oCurRun.Paragraph = oElement.Paragraph || oElement;
                                oNewRun = oCurRun.Copy2({CopyReviewPr : false});
                                oCurRun.Paragraph = oParentParagraph2;
                                if(oLastText.elements[oLastText.elements.length - 1] === oCurRun.Content[t])
                                {
                                    if(t < oCurRun.Content.length - 1)
                                    {
                                        oNewRun.Remove_FromContent(t + 1, oNewRun.Content.length - (t + 1), false);
                                    }
                                    aContentToInsert.splice(0, 0, oNewRun);
                                    break;
                                }
                            }
                        }
                        else
                        {
                            aContentToInsert.splice(0, 0, oCurRun.Copy(false, {CopyReviewPr : false}));
                        }
                        break;
                    }
                    else if(oLastText === oParentParagraph.Content[k])
                    {
                        aContentToInsert.splice(0, 0,  oParentParagraph.Content[k].Copy(false, {CopyReviewPr : false}));
                        break;
                    }
                }
                if( (oLastText.lastRun && oFirstText.firstRun) && oLastText.lastRun === oFirstText.firstRun || (!oLastText.lastRun && !oFirstText.firstRun) && oLastText === oFirstText)
                {
                    if(aContentToInsert.length > 0)
                    {
                        oNewRun = aContentToInsert[0];
                        if(oNewRun instanceof  ParaRun)
                        {
                            for(t = 0; t < oFirstText.firstRun.Content.length; ++t)
                            {
                                if(oFirstText.elements[0] === oFirstText.firstRun.Content[t])
                                {
                                    oNewRun.Remove_FromContent(0, t, false);
                                    break;
                                }
                            }
                        }
                    }
                }
                else
                {
                    for(k -= 1; k > -1; --k)
                    {
                        oCurRun = oParentParagraph.Content[k];
                        if(oCurRun !== oFirstRun && oCurRun !== oFirstText)
                        {
                            aContentToInsert.splice(0, 0, oCurRun.Copy(false, {CopyReviewPr : false}));
                        }
                        else
                        {
                            if(oCurRun === oFirstText)
                            {
                                aContentToInsert.splice(0, 0,  oCurRun.Copy(false, {CopyReviewPr : false}));
                            }
                            else
                            {

                                for(t = 0; t < oCurRun.Content.length; ++t)
                                {
                                    if(oFirstText.elements[0] === oCurRun.Content[t])
                                    {
                                        if(oLastText.lastRun === oFirstText.firstRun)
                                        {
                                            oNewRun = aContentToInsert[0];
                                        }
                                        else
                                        {
                                            oCurRun.Paragraph = oElement.Paragraph || oElement;
                                            oNewRun = oCurRun.Copy2({CopyReviewPr : false});
                                            oCurRun.Paragraph = oParentParagraph2;
                                            aContentToInsert.splice(0, 0, oNewRun);
                                        }
                                        oNewRun.Remove_FromContent(0, t, false);
                                    }
                                }
                            }
                            break;
                        }
                    }
                }

                if(oChange.remove.length === 0)
                {
                    if(aContentToInsert.length > 0)
                    {
                        var index = oChange.anchor.index;
                        oChildNode = oNode.children[index];
                        if(oChildNode)
                        {
                            oFirstText = oChildNode.element;
                            for(j = 0; j < oElement.Content.length; ++j)
                            {
                                if(Array.isArray(oElement.Content))
                                {
                                    oCurRun = oElement.Content[j];
                                    if(oFirstText === oCurRun)
                                    {
                                        for(t = aContentToInsert.length - 1; t > - 1; --t)
                                        {
                                            if(!(aContentToInsert[t].IsParaEndRun && aContentToInsert[t].IsParaEndRun()))
                                            {
                                                oElement.AddToContent(j + 1, aContentToInsert[t]);
                                            }
                                        }
                                        break;
                                    }
                                    else if(Array.isArray(oCurRun.Content))
                                    {
                                        for(k = 0; k < oCurRun.Content.length; ++k)
                                        {
                                            if(oFirstText.elements[0] === oCurRun.Content[k])
                                            {
                                                break;
                                            }
                                        }
                                        var bFind = false;
                                        if(k === oCurRun.Content.length)
                                        {
                                            if(oFirstText.firstRun === oCurRun)
                                            {
                                                k = 0;
                                                bFind = true;
                                            }
                                        }
                                        else
                                        {
                                            bFind = true;
                                        }
                                        if(k <= oCurRun.Content.length && bFind)
                                        {
                                            oCurRun.Split2(k, oElement, j);
                                            for(t = aContentToInsert.length - 1; t > - 1; --t)
                                            {
                                                if(!(aContentToInsert[t].IsParaEndRun && aContentToInsert[t].IsParaEndRun()))
                                                {
                                                    oElement.AddToContent(j + 1, aContentToInsert[t]);
                                                }
                                            }
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            //handle removed elements
            if(oChange.remove.length > 0)
            {
                oLastText = oChange.remove[oChange.remove.length - 1].element;
                oFirstText = oChange.remove[0].element;
                if(oLastText.lastRun)
                {
                    oCurRun = oLastText.lastRun;
                }
                else
                {
                    oCurRun = oLastText;
                }

                var nInsertPosition = -1;
                for(k = oElement.Content.length - 1; k > -1; --k)
                {
                    if(oElement.Content[k] === oCurRun)
                    {
                        if(oLastText instanceof CTextElement)
                        {
                            for(t = oCurRun.Content.length - 1; t > -1; t--)
                            {
                                if(oCurRun.Content[t] === oLastText.elements[oLastText.elements.length - 1])
                                {
                                    break;
                                }
                            }
                            if(t > -1)
                            {
                                //  if(t !== oCurRun.Content.length - 1)
                                {
                                    nInsertPosition = k + 1;
                                    oNewRun = oCurRun.Split2(t + 1, oElement, k);
                                    oNewRun.SetReviewTypeWithInfo(reviewtype_Common, oCurRun.ReviewInfo.Copy());
                                }
                                // else
                                // {}
                            }
                        }
                        else
                        {

                        }
                        break;
                    }
                }
                for(; k > -1; --k)
                {
                    oChildElement = oElement.Content[k];
                    if(oChildElement !== oFirstText.firstRun && oChildElement !== oFirstText)
                    {
                        if(oChildElement instanceof ParaRun)
                        {
                            this.updateReviewInfo(oChildElement, reviewtype_Remove);
                        }
                        else
                        {
                            if(oChildElement.Content)
                            {
                                this.setParagraphReviewInfoRecursive(oChildElement);
                            }
                        }
                    }
                    else
                    {
                        if(oChildElement instanceof ParaRun)
                        {
                            for(t = 0; t < oChildElement.Content.length; t++)
                            {
                                if(oChildElement.Content[t] === oFirstText.elements[0])
                                {
                                    break;
                                }
                            }
                            t = Math.min(Math.max(t, 0), oChildElement.Content.length - 1);
                            if(t > 0)
                            {
                                oNewRun = oChildElement.Split2(t, oElement, k);
                                this.updateReviewInfo(oNewRun, reviewtype_Remove);
                                nInsertPosition++;
                            }
                            else
                            {
                                this.updateReviewInfo(oChildElement, reviewtype_Remove);
                            }
                        }
                        else
                        {
                            if(oChildElement.Content)
                            {
                                this.setParagraphReviewInfoRecursive(oChildElement);
                            }
                        }
                        break;
                    }
                }
                if(nInsertPosition > -1)
                {
                    for(t = aContentToInsert.length - 1; t > - 1; --t)
                    {
                        if(!(aContentToInsert[t].IsParaEndRun && aContentToInsert[t].IsParaEndRun()))
                        {
                            oElement.AddToContent(nInsertPosition, aContentToInsert[t]);
                        }
                    }
                }
            }
        }
        for(i = 0; i < oNode.children.length; ++i)
        {
            oChildNode = oNode.children[i];
            if(Array.isArray(oChildNode.element.Content))
            {
                this.applyChangesToParagraph(oChildNode);
            }
            else
            {
                for(j = 0; j < oChildNode.children.length; ++j)
                {
                    if(oChildNode.children[j].element instanceof CDocumentContent)
                    {
                        this.applyChangesToDocContent(oChildNode.children[j]);
                    }
                }
            }
        }

        if(oNode.partner)
        {
            var oPartnerNode = oNode.partner;
            var oPartnerElement = oPartnerNode.element;
            if(oPartnerElement instanceof Paragraph)
            {
                var oOldPrChange = oPartnerElement.Pr.PrChange;
                oPartnerElement.Pr.PrChange = oElement.Pr;
                var oDiffPr = oPartnerElement.Pr.GetDiffPrChange(), oStyle;
                oPartnerElement.Pr.PrChange = oOldPrChange;

                oElement.Set_ContextualSpacing(oDiffPr.ContextualSpacing);

                if (oDiffPr.Ind)
                {
                    oElement.Set_Ind(oDiffPr.Ind, false);
                }

                if(undefined !== oDiffPr.Jc)
                {
                    oElement.Set_Align(oDiffPr.Jc);
                }
                if(undefined !== oDiffPr.KeepLines)
                {
                    oElement.Set_KeepLines(oDiffPr.KeepLines);
                }
                if(undefined !== oDiffPr.KeepNext)
                {
                    oElement.Set_KeepNext(oDiffPr.KeepNext);
                }
                if(undefined !== oDiffPr.PageBreakBefore)
                {
                    oElement.Set_PageBreakBefore(oDiffPr.PageBreakBefore);
                }

                if (oDiffPr.Spacing)
                    oElement.Set_Spacing(oDiffPr.Spacing, false);

                if (oDiffPr.Shd)
                    oElement.Set_Shd(oDiffPr.Shd, true);

//                        oElement.Set_WidowControl(oDiffPr.WidowControl);

                if (oDiffPr.Tabs)
                {
                    if(!oElement.Pr.Tabs || !oDiffPr.Tabs.Is_Equal(oElement.Pr.Tabs))
                    {
                        oElement.Set_Tabs(oDiffPr.Tabs);
                    }
                }

                if(oElement.Pr.NumPr && !oPartnerElement.Pr.NumPr)
                {
                    oElement.RemoveNumPr();
                }
                else
                {
                    if (oDiffPr.NumPr && this.revisedDocument.CopyNumberingMap[oDiffPr.NumPr.NumId])
                    {
                        if(oElement.Pr.NumPr && oPartnerElement.Pr.NumPr && oElement.Pr.NumPr.IsValid() && oPartnerElement.Pr.NumPr.IsValid())
                        {
                            if(oElement.Pr.NumPr.Lvl === oPartnerElement.Pr.NumPr.Lvl)
                            {
                                var oNumOrig = this.originalDocument.GetNumbering().GetNum(oElement.Pr.NumPr.NumId);
                                var oNumRevise = this.revisedDocument.GetNumbering().GetNum(oPartnerElement.Pr.NumPr.NumId);
                                if(oNumOrig && oNumRevise)
                                {
                                    if(!oNumOrig.IsSimilar(oNumRevise))
                                    {
                                        oElement.SetNumPr(this.revisedDocument.CopyNumberingMap[oDiffPr.NumPr.NumId], oDiffPr.NumPr.Lvl);
                                    }
                                }
                                else if(oNumOrig)
                                {
                                    oElement.RemoveNumPr();
                                }
                                else if(oNumRevise)
                                {
                                    oElement.SetNumPr(this.revisedDocument.CopyNumberingMap[oDiffPr.NumPr.NumId], oDiffPr.NumPr.Lvl);
                                }
                            }
                        }
                    }
                    else
                    {
                        if(oElement.Pr.NumPr)
                        {
                            oElement.RemoveNumPr();
                        }
                    }
                }
                if (oDiffPr.PStyle)
                {
                    if(oStyle = this.revisedDocument.Styles.Get(oDiffPr.PStyle))
                    {
                        var oStyleId = this.copyStyle(oStyle);
                        if(oStyleId !== oElement.Pr.PStyle)
                        {
                            oElement.Style_Add(oStyleId, true);
                        }
                    }
                }
                else
                {
                    if(oElement.Pr.PStyle && !oPartnerElement.Pr.PStyle)
                    {
                        oElement.Style_Add(undefined, true);
                    }
                }

                if (oDiffPr.Brd)
                    oElement.Set_Borders(oDiffPr.Brd);
            }
        }

    };

    CDocumentComparison.prototype.applyChangesToTable = function(oNode)
    {
        var oElement = oNode.element, oChange, i, j, oRow;
        oNode.changes.sort(function(c1, c2){return c2.anchor.index - c1.anchor.index});
        for(i = 0; i < oNode.changes.length; ++i)
        {
            oChange = oNode.changes[i];
            for(j = oChange.remove.length - 1; j > -1;  --j)
            {
                oRow = oChange.remove[j].element;
                for (var nCurCell = 0, nCellsCount = oRow.GetCellsCount(); nCurCell < nCellsCount; ++nCurCell)
                {
                    this.setDocContentReviewInfoRecursive(oRow.GetCell(nCurCell).GetContent());
                }
                oRow.SetReviewType(reviewtype_Remove);

            }
            for(j = oChange.insert.length - 1; j > -1;  --j)
            {
                oElement.Content.splice(oChange.anchor.index, 0, oChange.insert[j].element.Copy(oElement, {CopyReviewPr : false}));
                History.Add(new CChangesTableAddRow(oElement, oChange.anchor.index, [oElement.Content[oChange.anchor.index]]));
            }
            oElement.Internal_ReIndexing(0);
            if (oElement.Content.length > 0 && oElement.Content[0].Get_CellsCount() > 0)
                oElement.CurCell = oElement.Content[0].Get_Cell(0);
        }
        for(i = 0; i < oNode.children.length; ++i)
        {
            this.applyChangesToTableRow(oNode.children[i]);
        }
    };

    CDocumentComparison.prototype.applyChangesToTableRow = function(oNode)
    {
        //TODO: handle cell inserts and removes

        for(var i = 0; i < oNode.children.length; ++i)
        {
            this.applyChangesToDocContent(oNode.children[i]);
        }
    };


    CDocumentComparison.prototype.copyTableStylePr = function(oPr)
    {
        var oCopyPr = oPr.Copy();
        if(undefined !== oCopyPr.ParaPr.NumPr && undefined !== oCopyPr.ParaPr.NumPr.NumId)
        {
            var NewId = this.revisedDocument.CopyNumberingMap[oCopyPr.ParaPr.NumPr.NumId];
            if (undefined !== NewId)
                oCopyPr.ParaPr.SetNumPr(NewId, oCopyPr.ParaPr.NumPr.Lvl);
        }
        return oCopyPr;
    };

    CDocumentComparison.prototype.copyStyle = function(oStyle)
    {
        if(!oStyle)
        {
            return null;
        }
        if(this.StylesMap[oStyle.Id])
        {
            return this.StylesMap[oStyle.Id];
        }
        var oStyleCopy;
        if(oStyleCopy = this.originalDocument.Styles.GetStyleIdByName(oStyle.Name, false))
        {
            return oStyleCopy;
        }
        oStyleCopy = oStyle.Copy();
        oStyleCopy.Set_Name(oStyle.Name);
        oStyleCopy.Set_Next(oStyle.Next);
        oStyleCopy.Set_Type(oStyle.Type);
        oStyleCopy.Set_QFormat(oStyle.qFormat);
        oStyleCopy.Set_UiPriority(oStyle.uiPriority);
        oStyleCopy.Set_Hidden(oStyle.hidden);
        oStyleCopy.Set_SemiHidden(oStyle.semiHidden);
        oStyleCopy.Set_UnhideWhenUsed(oStyle.unhideWhenUsed);
        oStyleCopy.Set_TextPr(oStyle.TextPr.Copy());
        var oParaPrCopy = oStyle.ParaPr.Copy();

        if(undefined !== oParaPrCopy.NumPr && undefined !== oParaPrCopy.NumPr.NumId)
        {
            var NewId = this.revisedDocument.CopyNumberingMap[oParaPrCopy.NumPr.NumId];
            if (undefined !== NewId)
                oParaPrCopy.SetNumPr(NewId, oParaPrCopy.NumPr.Lvl);
        }

        oStyleCopy.Set_ParaPr(oParaPrCopy);
        oStyleCopy.Set_TablePr(oStyle.TablePr.Copy());
        oStyleCopy.Set_TableRowPr(oStyle.TableRowPr.Copy());
        oStyleCopy.Set_TableCellPr(oStyle.TableCellPr.Copy());
        if (undefined !== oStyle.TableBand1Horz)
        {
            oStyleCopy.Set_TableBand1Horz(this.copyTableStylePr(oStyle.TableBand1Horz));
            oStyleCopy.Set_TableBand1Vert(this.copyTableStylePr(oStyle.TableBand1Vert));
            oStyleCopy.Set_TableBand2Horz(this.copyTableStylePr(oStyle.TableBand2Horz));
            oStyleCopy.Set_TableBand2Vert(this.copyTableStylePr(oStyle.TableBand2Vert));
            oStyleCopy.Set_TableFirstCol(this.copyTableStylePr(oStyle.TableFirstCol));
            oStyleCopy.Set_TableFirstRow(this.copyTableStylePr(oStyle.TableFirstRow));
            oStyleCopy.Set_TableLastCol(this.copyTableStylePr(oStyle.TableLastCol));
            oStyleCopy.Set_TableLastRow(this.copyTableStylePr(oStyle.TableLastRow));
            oStyleCopy.Set_TableTLCell(this.copyTableStylePr(oStyle.TableTLCell));
            oStyleCopy.Set_TableTRCell(this.copyTableStylePr(oStyle.TableTRCell));
            oStyleCopy.Set_TableBLCell(this.copyTableStylePr(oStyle.TableBLCell));
            oStyleCopy.Set_TableBRCell(this.copyTableStylePr(oStyle.TableBRCell));
            oStyleCopy.Set_TableWholeTable(this.copyTableStylePr(oStyle.TableWholeTable));
        }


        if(oStyle.BasedOn)
        {
            if(!this.StylesMap[oStyle.BasedOn])
            {
                oStyleCopy.Set_BasedOn(this.copyStyle(this.revisedDocument.Styles.Get(oStyle.BasedOn)));
            }
            else
            {
                oStyleCopy.Set_BasedOn(this.StylesMap[oStyle.BasedOn])
            }
        }
        this.originalDocument.Styles.Add(oStyleCopy);
        this.StylesMap[oStyleCopy.Id] = oStyleCopy.Id;
        return oStyleCopy.Id;
    };


    CDocumentComparison.prototype.replaceParagraphStyle = function(oParagraph)
    {
        var oStyle;
        if(oParagraph.Pr && oParagraph.Pr.PStyle && (oStyle = this.revisedDocument.Styles.Get(oParagraph.Pr.PStyle)))
        {
            oParagraph.Style_Add(this.copyStyle(oStyle), true);
        }

        if (oParagraph.Pr.NumPr && oParagraph.Pr.NumPr.NumId && this.revisedDocument.CopyNumberingMap[oParagraph.Pr.NumPr.NumId])
            oParagraph.SetNumPr(this.revisedDocument.CopyNumberingMap[oParagraph.Pr.NumPr.NumId], oParagraph.Pr.NumPr.Lvl);
    };

    CDocumentComparison.prototype.replaceTableStyle = function(oTable)
    {
        var oStyle;
        if(oTable.Pr && oTable.Pr.PStyle && (oStyle = this.revisedDocument.Styles.Get(oTable.TableStyle)))
        {
            oTable.Set_TableStyle(this.copyStyle(oStyle), false);
        }
    };

    CDocumentComparison.prototype.replaceTableStyles = function(oTable)
    {
        for(var i = 0; i < oTable.Content.length; ++i)
        {
            var oRow = oTable.Content[i];
            for(var j = 0; j < oRow.Content.length; ++j)
            {
                this.replaceDocContentStyles(oRow.Content[j].Content);
            }
        }
        this.replaceTableStyle(oTable);
    };

    CDocumentComparison.prototype.replaceDocContentStyles = function(oContent)
    {
        for(var i = 0; i < oContent.Content.length; ++i)
        {
            this.replaceElementStyles(oContent.Content[i]);
        }
    };

    CDocumentComparison.prototype.replaceElementStyles = function(oChildElement)
    {
        if(oChildElement instanceof Paragraph)
        {
            this.replaceParagraphStyle(oChildElement)
        }
        else if(oChildElement.Content instanceof CDocumentContent)
        {
            this.replaceDocContentStyles(oChildElement.Content);
        }
        else if(oChildElement instanceof CTable)
        {
            this.replaceTableStyles(oChildElement);
        }
    };

    CDocumentComparison.prototype.setElementReviewInfoRecursive = function(oChildElement)
    {
        if(oChildElement instanceof Paragraph)
        {
            this.setParagraphReviewInfoRecursive(oChildElement);
        }
        else if(oChildElement instanceof CDocumentContent)
        {
            this.setDocContentReviewInfoRecursive(oChildElement);
        }
        else if(oChildElement instanceof CTable)
        {
            this.setTableReviewInfoRecursive(oChildElement);
        }
    };

    CDocumentComparison.prototype.insertNodesToDocContent = function(oElement, nIndex, aInsert)
    {

        var k = 0;
        for(var j = 0; j < aInsert.length; ++j)
        {
            var oChildElement = null;
            if(aInsert[j].element.Get_Type)
            {
                oChildElement = aInsert[j].element.Copy(oElement, oElement.DrawingDocument,  {CopyReviewPr : false});
                this.replaceParagraphStyle(oChildElement);
            }
            else
            {
                if(aInsert[j].element.Parent && aInsert[j].element.Parent.Get_Type)
                {
                    oChildElement = aInsert[j].element.Parent.Copy(oElement, oElement.DrawingDocument,  {CopyReviewPr : false});
                }
            }
            if(oChildElement)
            {
                this.replaceElementStyles(oChildElement);
                oElement.Internal_Content_Add(nIndex + k, oChildElement, false);
                ++k;
            }
        }
    };

    CDocumentComparison.prototype.applyChangesToChildNode = function(oChildNode)
    {
        var oChildElement = oChildNode.element;
        if(oChildElement instanceof Paragraph)
        {
            this.applyChangesToParagraph(oChildNode);
        }
        else if(oChildElement instanceof CDocumentContent)
        {
            this.applyChangesToDocContent(oChildNode);
        }
        else if(oChildElement instanceof CTable)
        {
            this.applyChangesToTable(oChildNode);
        }
    };

    CDocumentComparison.prototype.applyChangesToDocContent = function(oNode)
    {
        var oElement = oNode.element, oChange, i, j, k, oChildElement, oChildNode, oPartnerNode, oPartnerElement, oOldPrChange, oDiffPr, oStyle;

        oNode.changes.sort(function(c1, c2){return c2.anchor.index - c1.anchor.index});
        for(i = 0; i < oNode.changes.length; ++i)
        {
            oChange = oNode.changes[i];
            for(j = oChange.remove.length - 1; j > -1; --j)
            {
                oChildNode = oChange.remove[j];
                oChildElement = oChildNode.element;
                this.setElementReviewInfoRecursive(oChildElement);
            }
            this.insertNodesToDocContent(oElement, oChange.anchor.index + oChange.remove.length, oChange.insert);
        }
        for(i = 0; i < oNode.children.length; ++i)
        {
            this.applyChangesToChildNode(oNode.children[i]);
        }
    };

    CDocumentComparison.prototype.updateReviewInfo = function(oObject, nType, bParaEnd)
    {
        if(!bParaEnd && oObject.IsParaEndRun &&oObject.IsParaEndRun() )
        {
            return;
        }
        if(oObject.ReviewInfo && oObject.ReviewType === reviewtype_Common)
        {
            var oCore = this.revisedDocument.Core;
            var oReviewIno = oObject.ReviewInfo.Copy();
            oReviewIno.Editor   = this.api;
            oReviewIno.UserId   = "";
            oReviewIno.MoveType = Asc.c_oAscRevisionsMove.NoMove;
            oReviewIno.PrevType = -1;
            oReviewIno.PrevInfo = null;
            oReviewIno.UserName = this.getUserName();
            if(oCore)
            {
                if(oCore.modified instanceof Date)
                {
                    oReviewIno.DateTime = oCore.modified.getTime();
                }
            }
            else
            {
                oReviewIno.DateTime = "Unknown";
            }
            oObject.SetReviewTypeWithInfo(nType, oReviewIno, false);
            if(nType === reviewtype_Remove && oObject.CheckRunContent)
            {
                var oThis = this;
                oObject.CheckRunContent(function (oRun) {
                    if(Array.isArray(oRun.Content))
                    {
                        for(var i = 0; i < oRun.Content.length; ++i)
                        {
                            if(oRun.Content[i].Type === para_Drawing)
                            {
                                var aContents = oRun.Content[i].GetAllDocContents();
                                for(var j = 0; j < aContents.length; ++j)
                                {
                                    oThis.setDocContentReviewInfoRecursive(aContents[j]);
                                }
                            }
                            else if(oRun.Content[i].Type === para_FootnoteReference)
                            {
                                oThis.setDocContentReviewInfoRecursive(oRun.Content[i].Footnote);
                            }
                        }
                    }
                });
            }
        }
    };

    CDocumentComparison.prototype.createNodeFromDocContent = function(oElement, oParentNode, oHashWords)
    {
        var oRet = new CNode(oElement, oParentNode);
        var bRoot = (oParentNode === null);
        for(var i = 0; i < oElement.Content.length; ++i)
        {
            var oChElement = oElement.Content[i];
            if(oChElement instanceof Paragraph)
            {
                if(bRoot)
                {
                    oHashWords = new Minhash({});
                }
                var oParagraphNode = this.createNodeFromRunContentElement(oChElement, oRet, oHashWords);
                if(bRoot)
                {
                    oParagraphNode.hashWords = oHashWords;
                }
            }
            else if(oChElement instanceof CBlockLevelSdt)
            {
                if(bRoot)
                {
                    oHashWords = new Minhash({});
                }
                var oBlockNode = this.createNodeFromDocContent(oChElement.Content, oRet, oHashWords);
                if(bRoot)
                {
                    oBlockNode.hashWords = oHashWords;
                }
            }
            else if(oChElement instanceof CTable)
            {
                if(this.options.getTables())
                {
                    if(bRoot)
                    {
                        oHashWords = new Minhash({});
                    }
                    var oTableNode = new CNode(oChElement, oRet);
                    if(bRoot)
                    {
                        oHashWords = new Minhash({});
                        oTableNode.hashWords = oHashWords;
                    }
                    for(var j = 0; j < oChElement.Content.length; ++j)
                    {
                        var oRowNode = new CNode(oChElement.Content[j], oTableNode);
                        for(var k = 0; k < oChElement.Content[j].Content.length; ++k)
                        {
                            this.createNodeFromDocContent(oChElement.Content[j].Content[k].Content, oRowNode, oHashWords);
                        }
                    }
                }
            }
            else
            {
                var oNode = new CNode(oChElement, oRet);
                if(bRoot)
                {
                    oHashWords = new Minhash({});
                    oNode.hashWords = oHashWords;
                }
            }

        }
        return oRet;
    };

    CDocumentComparison.prototype.createNodeFromRunContentElement = function(oElement, oParentNode, oHashWords)
    {
        var oRet = new CNode(oElement, oParentNode);
        var oLastText = null, oRun, oRunElement, i, j;
        var aLastWord = [];
        for(i = 0; i < oElement.Content.length; ++i)
        {
            oRun = oElement.Content[i];
            if(oRun instanceof ParaRun)
            {
                if(oRun.Content.length > 0)
                {
                    if(this.options.getWords())
                    {
                        if(!oLastText)
                        {
                            oLastText = new CTextElement();
                            oLastText.setFirstRun(oRun);
                        }
                        if(oLastText.elements.length === 0)
                        {
                            oLastText.setFirstRun(oRun);
                            oLastText.setLastRun(oRun);
                        }
                        for(j = 0; j < oRun.Content.length; ++j)
                        {
                            oRunElement = oRun.Content[j];
                            var bPunctuation = !!(para_Text === oRunElement.Type && (AscCommon.g_aPunctuation[oRunElement.Value] || EXCLUDED_PUNCTUATION[oRunElement.Value]));
                            if(oRunElement.Type === para_Space || oRunElement.Type === para_Tab
                                || oRunElement.Type === para_Separator || oRunElement.Type === para_NewLine
                                || oRunElement.Type === para_FootnoteReference
                                || bPunctuation)
                            {
                                if(bPunctuation)
                                {
                                    if(EXCLUDED_PUNCTUATION[oRunElement.Value])
                                    {
                                        bPunctuation = false;
                                    }
                                }
                                if(oLastText.elements.length > 0 && (!bPunctuation || (AscCommon.g_aPunctuation[oRunElement.Value] & AscCommon.PUNCTUATION_FLAG_CANT_BE_AT_END) ))
                                {
                                    new CNode(oLastText, oRet);
                                    oLastText.updateHash(oHashWords);
                                    oLastText = new CTextElement();
                                    oLastText.setFirstRun(oRun);
                                }

                                oLastText.setLastRun(oRun);
                                oLastText.elements.push(oRunElement);

                                if (!bPunctuation || (AscCommon.g_aPunctuation[oRunElement.Value] & AscCommon.PUNCTUATION_FLAG_CANT_BE_AT_BEGIN) )
                                {
                                    oLastText.updateHash(oHashWords);
                                    new CNode(oLastText, oRet);
                                    oLastText = new CTextElement();
                                    oLastText.setFirstRun(oRun);
                                    oLastText.setLastRun(oRun);
                                }
                            }
                            else if(oRunElement.Type === para_Drawing)
                            {
                                if(oLastText.elements.length > 0)
                                {
                                    oLastText.updateHash(oHashWords);
                                    new CNode(oLastText, oRet);
                                    oLastText = new CTextElement();
                                    oLastText.setFirstRun(oRun);
                                    oLastText.setLastRun(oRun);
                                }
                                oLastText.elements.push(oRun.Content[j]);
                                new CNode(oLastText, oRet);
                                oLastText = new CTextElement();
                                oLastText.setFirstRun(oRun);
                                oLastText.setLastRun(oRun);
                            }
                            else if(oRunElement.Type === para_End)
                            {
                                if(oLastText.elements.length > 0)
                                {
                                    oLastText.updateHash(oHashWords);
                                    new CNode(oLastText, oRet);
                                    oLastText = new CTextElement();
                                    oLastText.setFirstRun(oRun);
                                    oLastText.setLastRun(oRun);
                                }
                                oLastText.setFirstRun(oRun);
                                oLastText.setLastRun(oRun);
                                oLastText.elements.push(oRun.Content[j]);
                                new CNode(oLastText, oRet);
                                //oLastText.updateHash(oHashWords);
                                oLastText = new CTextElement();
                                oLastText.setFirstRun(oRun);
                                oLastText.setLastRun(oRun);
                            }
                            else
                            {
                                if(oLastText.elements.length === 0)
                                {
                                    oLastText.setFirstRun(oRun);
                                }
                                oLastText.setLastRun(oRun);
                                oLastText.elements.push(oRun.Content[j]);
                            }
                        }
                    }
                    else
                    {
                        if(oLastText && oLastText.elements.length > 0)
                        {
                            new CNode(oLastText, oRet);
                        }
                        for(j = 0; j < oRun.Content.length; ++j)
                        {
                            oRunElement = oRun.Content[j];
                            if(AscFormat.isRealNumber(oRunElement.Value))
                            {
                                aLastWord.push(oRunElement.Value);
                            }
                            else
                            {
                                if(aLastWord.length > 0)
                                {
                                    oHashWords.update(aLastWord);
                                    aLastWord.length = 0;
                                }
                            }
                            oLastText = new CTextElement();
                            oLastText.setFirstRun(oRun);
                            oLastText.setLastRun(oRun);
                            oLastText.elements.push(oRunElement);
                            new CNode(oLastText, oRet);
                        }
                        oLastText = new CTextElement();
                        oLastText.setFirstRun(oRun);
                        oLastText.setLastRun(oRun);
                    }
                }
            }
            else
            {
                if(Array.isArray(oRun.Content))
                {
                    if(oLastText && oLastText.elements.length > 0)
                    {
                        new CNode(oLastText, oRet);
                    }
                    if(aLastWord.length > 0)
                    {
                        oHashWords.update(aLastWord);
                        aLastWord.length = 0;
                    }
                    oLastText = null;
                    this.createNodeFromRunContentElement(oRun, oRet, oHashWords);
                }
            }
        }
        if(oLastText && oLastText.elements.length > 0)
        {
            oLastText.updateHash(oHashWords);
            new CNode(oLastText, oRet);
        }
        return oRet;
    };



    window['AscCommonWord'] = window['AscCommonWord'] || {};
    window['AscCommonWord'].CDocumentComparison = CDocumentComparison;
    window['AscCommonWord'].ComparisonOptions = window['AscCommonWord']["ComparisonOptions"] = ComparisonOptions;

    function CompareBinary(oApi, sBinary2, oOptions)
    {

        var oDoc1 = oApi.WordControl.m_oLogicDocument;

        oApi.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.SlowOperation);
        var bHaveRevisons2 = false;
        var oDoc2 = AscFormat.ExecuteNoHistory(function(){
            var oTableId =  AscCommon.g_oTableId;
            AscCommon.g_oTableId = new AscCommon.CTableId();
            AscCommon.g_oTableId.init();
            AscCommon.g_oIdCounter.m_bLoad = true;
            var oBinaryFileReader, openParams        = {checkFileSize : /*this.isMobileVersion*/false, charCount : 0, parCount : 0};
            var oDoc2 = new CDocument(oApi.WordControl.m_oDrawingDocument, true);
            oDoc2.Footnotes = oDoc1.Footnotes;
            oApi.WordControl.m_oDrawingDocument.m_oLogicDocument = oDoc2;
            oDoc2.ForceCopySectPr = true;
            oBinaryFileReader = new AscCommonWord.BinaryFileReader(oDoc2, openParams);

            oApi.WordControl.m_oLogicDocument = oDoc2;
            AscCommon.pptx_content_loader.Start_UseFullUrl(oApi.insertDocumentUrlsData);
            if (!oBinaryFileReader.Read(sBinary2))
            {
                oDoc2 = null;
            }
            if(oDoc2)
            {
                oDoc2.ForceCopySectPr = false;

                bHaveRevisons2 = oDoc2.HaveRevisionChanges(false);
                oDoc2.Start_SilentMode();

                var LogicDocuments = oDoc2.TrackRevisionsManager.Get_AllChangesLogicDocuments();
                for (var LogicDocId in LogicDocuments)
                {
                    var LogicDoc = AscCommon.g_oTableId.Get_ById(LogicDocId);
                    if (LogicDoc)
                    {
                        LogicDoc.AcceptRevisionChanges(undefined, true);
                    }
                }
                oDoc2.End_SilentMode(false);
            }

            AscCommon.g_oIdCounter.m_bLoad = false;
            oApi.WordControl.m_oDrawingDocument.m_oLogicDocument = oDoc1;
            oApi.WordControl.m_oLogicDocument = oDoc1;

            AscCommon.g_oTableId = oTableId;
            return oDoc2;
        }, this, []);
        oDoc1.History.Document = oDoc1;
        if(oDoc2)
        {
            var fCallback = function()
            {
                var oComp = new AscCommonWord.CDocumentComparison(oDoc1, oDoc2, oOptions ? oOptions : new ComparisonOptions());
                oComp.compare();
            };
            if(!window['NATIVE_EDITOR_ENJINE'] && (oDoc1.HaveRevisionChanges(false) || bHaveRevisons2))
            {

                oApi.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.SlowOperation);
                oApi.sendEvent("asc_onAcceptChangesBeforeCompare", function (bAccept) {
                    if(bAccept){
                        oApi.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.SlowOperation);
                        fCallback();
                    }
                    else {
                    }
                })
            }
            else
            {

                fCallback();
            }
        }
        else
        {
            AscCommon.pptx_content_loader.End_UseFullUrl();
        }
    }
    window['AscCommonWord']["CompareBinary"] =  window['AscCommonWord'].CompareBinary = CompareBinary;
    window['AscCommonWord']["ComparisonOptions"] = window['AscCommonWord'].ComparisonOptions = ComparisonOptions;
})();

