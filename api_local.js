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
    window['Asc']['Addons'] = window['Asc']['Addons'] || {};
	window['Asc']['Addons']['comparison'] = true; // register addon
    window['Asc']['asc_docs_api'].prototype["asc_CompareDocumentUrl"] = window['Asc']['asc_docs_api'].prototype.asc_CompareDocumentUrl = function (sUrl, oOptions, token) {
        var t = this;
        t.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Waiting);
        window["AscDesktopEditor"]["DownloadFiles"]([sUrl], [], function(files) {
            t.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Waiting);
            if (Array.isArray(files) && files.length == 1)
            {
                return window["AscDesktopEditor"]["CompareDocumentUrl"](files[0], oOptions);
            }
        });
    };
    window['Asc']['asc_docs_api'].prototype["asc_CompareDocumentFile"] = window['Asc']['asc_docs_api'].prototype.asc_CompareDocumentFile = function (oOptions) {
        window["AscDesktopEditor"]["OpenFilenameDialog"]("word", false, function(_file){
            var file = _file;
            if (Array.isArray(file))
                file = file[0];
            
            if (file && "" != file)
                window["AscDesktopEditor"]["CompareDocumentFile"](file, oOptions);
        });
    };
    // for local files in desktop apps
    window["onDocumentCompare"] = function(folder, file_content, file_content_len, image_map, options) {
        var api = window.editor;
        if (file_content == "")
        {
            api.sendEvent("asc_onError", Asc.c_oAscError.ID.ConvertationOpenError, Asc.c_oAscError.Level.NoCritical);
            return;
        }

        var file = {
            IsValid : function() { return true; },
            GetBinary: function() { return AscCommon.getBinaryArray(file_content, file_content_len); },
            GetImageMap: function() { return image_map; }
        };

        AscCommonWord.CompareDocuments(api, file);
    };
})(window, window.document);