"""Exact bookmark-endpoint adapter for experimental fragments only."""

import optimized_pagination as v1


_BOOKMARK_RANGE = r'''
property bookmarkRecords : {}

on bookmarkPosition(d, markName)
    tell application "Microsoft Word"
        set matchedRecord to missing value
        repeat with bookmarkRecord in my bookmarkRecords
            set b to contents of bookmarkRecord
            if name of b is markName then
                if matchedRecord is not missing value then error "Bookmark name duplicate: " & markName
                set matchedRecord to b
            end if
        end repeat
        if matchedRecord is missing value then error "Bookmark absent: " & markName
        if story type of matchedRecord is not main text story then error "Bookmark outside main story: " & markName
        set s to start of bookmark of matchedRecord
        set e to end of bookmark of matchedRecord
        if s is missing value or e is missing value then error "Bookmark endpoints missing: " & markName
        if class of s is not integer or class of e is not integer then error "Bookmark endpoints invalid: " & markName
        if s < 0 or s is not e then error "Bookmark is not a collapsed measurement: " & markName
        return s
    end tell
end bookmarkPosition

on bookmarkRange(d, markName)
    set s to my bookmarkPosition(d, markName)
    tell application "Microsoft Word"
        set r to create range d start s end s
        if r is missing value then error "Created range missing: " & markName
        if start of content of r is not s or end of content of r is not s then error "Created range endpoints mismatch: " & markName
        return r
    end tell
end bookmarkRange
'''


def _snapshot_script(script):
    return script.replace(
        'if full name of d is not expectedName then error "Document path mismatch."',
        '''if full name of d is not expectedName then error "Document path mismatch."
        set my bookmarkRecords to get properties of bookmarks of d''',
    )


def _page_script(script):
    return _BOOKMARK_RANGE + _snapshot_script(script).replace(
        "set r to text object of bookmark (contents of bookmarkName) of d",
        "try\n            set r to my bookmarkRange(d, contents of bookmarkName)",
    ).replace(
        "        end repeat",
        '''            on error messageText number errorNumber
                error "Bookmark " & (contents of bookmarkName) & ": " & messageText number errorNumber
            end try
        end repeat''',
    )


_OBSERVATION_SCRIPT = _BOOKMARK_RANGE + _snapshot_script(v1.OBSERVATION_SCRIPT).replace(
    "set expectedText to item (i + 1) of argv",
    "try\n            set expectedText to item (i + 1) of argv",
).replace(
    'start of content of text object of bookmark (markName & "_s") of d',
    'my bookmarkPosition(d, markName & "_s")',
).replace(
    'start of content of text object of bookmark (markName & "_e") of d',
    'my bookmarkPosition(d, markName & "_e")',
).replace(
    "        end repeat\n    end tell",
    '''            on error messageText number errorNumber
                error "Observation bookmarks " & markName & "_s / " & markName & "_e: " & messageText number errorNumber
            end try
        end repeat
    end tell''',
)

_SCRIPTS = {
    v1.PAGE_SCRIPT: _page_script(v1.PAGE_SCRIPT),
    v1.POSITION_SCRIPT: _page_script(v1.POSITION_SCRIPT),
    v1.OBSERVATION_SCRIPT: _OBSERVATION_SCRIPT,
}


class FragmentNativePass(v1.NativePass):
    """Retain V1 algorithms, parsing and deadlines without its bookmark Range getter."""

    def script(self, phase, script, *arguments, limit=60):
        return super().script(phase, _SCRIPTS.get(script, script), *arguments, limit=limit)
