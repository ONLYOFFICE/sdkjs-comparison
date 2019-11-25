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


(function(window, document) {
    window['Asc']['asc_docs_api'].prototype["asc_CompareDocumentUrl"] = window['Asc']['asc_docs_api'].prototype.asc_CompareDocumentUrl = function (sUrl, oOptions, token) {
        this._CompareDocument({url: sUrl, format: "docx", token: token}, oOptions);
    };
    window['Asc']['asc_docs_api'].prototype["asc_CompareDocumentFile"] = window['Asc']['asc_docs_api'].prototype.asc_CompareDocumentFile = function (oOptions) {
        var t = this;
        AscCommon.ShowDocumentFileDialog(function (error, files) {
            if (c_oAscError.ID.No !== error) {
                t.sendEvent("asc_onError", error, c_oAscError.Level.NoCritical);
                return;
            }
            var format = AscCommon.GetFileExtension(files[0].name);
            var reader = new FileReader();
            reader.onload = function () {
                t._CompareDocument({data: new Uint8Array(reader.result), format: format}, oOptions);
            };
            reader.onerror = function () {
                t.sendEvent("asc_onError", c_oAscError.ID.Unknown, c_oAscError.Level.NoCritical);
            };
            reader.readAsArrayBuffer(files[0]);
        });
    };
    window['Asc']['asc_docs_api'].prototype._CompareDocument = function (document, oOptions) {
        var stream = null;
        var oApi = this;
        this.insertDocumentUrlsData = {
            imageMap: null, documents: [document], convertCallback: function (_api, url) {
                _api.insertDocumentUrlsData.imageMap = url;
                if (!url['output.bin']) {
                    _api.endInsertDocumentUrls();
                    _api.sendEvent("asc_onError", c_oAscError.ID.MailMergeLoadFile,
                        c_oAscError.Level.NoCritical);
                    return;
                }
                AscCommon.loadFileContent(url['output.bin'], function (httpRequest) {
                    if (null === httpRequest || !(stream = AscCommon.initStreamFromResponse(httpRequest))) {
                        _api.endInsertDocumentUrls();
                        _api.sendEvent("asc_onError", c_oAscError.ID.MailMergeLoadFile,
                            c_oAscError.Level.NoCritical);
                        return;
                    }
                    _api.endInsertDocumentUrls();
                }, "arraybuffer");
            }, endCallback: function (_api) {

                if (stream) {
                    AscCommonWord.CompareBinary(oApi, stream, oOptions);
                    stream = null;
                }
            }
        };

        var options = new Asc.asc_CDownloadOptions(Asc.c_oAscFileType.CANVAS_WORD);
        options.isNaturalDownload = true;
        this.asc_DownloadAs(options);
    };
})(window, window.document);
