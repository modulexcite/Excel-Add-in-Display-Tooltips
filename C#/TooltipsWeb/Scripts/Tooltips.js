/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/
// This function is run when the app is ready to start 
// interacting with the host application. It ensures 
// the DOM is ready before running the rest of the code.
Office.initialize = function (reason) {
    $(document).ready(function () {
        $('#span_help_01').hide();
        $('#span_help_02').hide();
        // Get the position of the keywords.
        var tooltip01Position = $('#keyword01').offset();
        var tooltip02Position = $('#keyword02').offset();
        $('#keyword01').mouseover(function () {
            // Show the tool tip at the specified position
            // and then slowly fade away.
           $('#span_help_01').show();
           $('#span_help_01').css({
            'position': 'absolute',
            'left': tooltip01Position.left,
            'top': tooltip01Position.top + 20
           }).fadeOut(4500);
        }); //end of keyword01 mouseover
        $('#keyword02').mouseover(function() {
            $('#span_help_02').show();
            $('#span_help_02').css({
                'position': 'absolute',
                'left': tooltip02Position.left,
                'top': tooltip02Position.top + 20
            }).fadeOut(4500);
        }); //end of keyword02 mouseover
    });
};
// *********************************************************
//
// Excel-Add-in-Display-Tooltips, https://github.com/OfficeDev/Excel-Add-in-Display-Tooltips
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************

