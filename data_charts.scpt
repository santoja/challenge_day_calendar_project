tell application "Microsoft Excel"
	
	tell worksheet 1 of workbook 1
		set theList to {{"categorie", "week -4", "week -3", "week -2", "week -1", "week"}, {"meetings", 10, 5, 9, 14, 20}, {"review", 9, 10, 5, 20, 6}, {"tickets", 7, 4, 5, 10.4, 17}, {"1:1", 3, 2, 3, 4, 5}}
		set listSize to count of theList
		
		set myRange to range ("A1:F" & listSize)
		
		set value of myRange to theList
		
		set value of cell "G1" to "Total"
		
		repeat with counter from 2 to listSize
			set value of cell ("G" & counter) to ("=SUM(B" & counter & ":F" & counter & ")")
		end repeat
		
		set objChart1 to make new chart object at end with properties {left position:530, top:1, height:300, width:500, name:"MyChart"}
		set ochart1 to chart of objChart1
		tell ochart1
			set newSeries to make new series at end with properties {series values:myRange}
			set series values of newSeries to myRange
			set has title to true
			tell its chart title
				set caption to "Category x Week"
			end tell
		end tell
		
		
		set objChart3 to make new chart object at end with properties {left position:530, top:320, height:300, width:500, name:"MyChart4"}
		set ochart3 to chart of objChart3
		tell ochart3
			set newSeries to make new series at end with properties {series values:myRange}
			set series values of newSeries to myRange
			set has title to true
			set chart type to line markers
			tell its chart title
				set caption to "Category x Week"
			end tell
		end tell
		
		set objChart2 to make new chart object at end with properties {left position:1050, top:1, height:300, width:500, name:"MyChart2"}
		set ochart2 to chart of objChart2
		tell ochart2
			set newSeries to make new series at end with properties {series values:myRange}
			set series values of newSeries to myRange
			set has title to true
			set plot by to by rows
			tell its chart title
				set caption to "Week x Category"
			end tell
		end tell
		
		set objChart3 to make new chart object at end with properties {left position:1050, top:320, height:300, width:500, name:"MyChart3"}
		set ochart3 to chart of objChart3
		tell ochart3
			set newSeries to make new series at end with properties {series values:myRange}
			set series values of newSeries to myRange
			set has title to true
			set chart type to line markers
			set plot by to by rows
			tell its chart title
				set caption to "Week x Category"
			end tell
		end tell
		
		set myRangeTotal to range ("G2:G" & listSize)
		
		set objChart5 to make new chart object at end with properties {left position:10, top:320, height:300, width:500, name:"MyChart5"}
		set ochart5 to chart of objChart5
		tell ochart5
			set newSeries to make new series at end with properties {series values:myRangeTotal}
			set has title to true
			set has legend to false
			set chart type to doughnut
			tell its chart title
				set caption to "Total"
			end tell
		end tell
		
		
	end tell
	
end tell
