/**
 * Returns all RegExMatch results in an array: [RegExMatchInfo1, RegExMatchInfo2, ...]
 * @param Haystack The string whose content is searched.
 * @param NeedleRegEx The RegEx pattern to search for.
 * @param StartingPosition If StartingPos is omitted, it defaults to 1 (the beginning of Haystack).
 * @returns {Array}
 */
RegExMatchAll(Haystack, NeedleRegEx, StartingPosition := 1) {
    out := []
	While StartingPosition := RegExMatch(Haystack, NeedleRegEx, &OutputVar, StartingPosition) {
		out.Push(OutputVar), StartingPosition += OutputVar[0] ? StrLen(OutputVar[0]) : 1
	}
	return out
}