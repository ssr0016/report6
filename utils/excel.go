package utils

import (
	"fmt"
	"reports/data/response"
	"strconv"
	"strings"

	"github.com/tealeg/xlsx"
)

func AddReportToSheet(sheet *xlsx.Sheet, report *response.ReportResponse) {
	// Set specific column widths
	const (
		orgNameWidth = 35
		titleWidth   = 25
	)

	// Add organization name
	orgNameRow := sheet.AddRow()
	orgNameCell := orgNameRow.AddCell()
	orgNameCell.Value = "ANG MANANAMPALATAYANG GUMAWA"
	orgNameCell.SetStyle(GetOrgNameStyle())
	orgNameCell.HMerge = 10

	// Set width for organization name row
	for i := 0; i <= 10; i++ {
		sheet.Col(i).Width = orgNameWidth
	}

	// Add main title
	titleRow := sheet.AddRow()
	titleCell := titleRow.AddCell()
	titleCell.Value = "NATIONAL WORKERS' MONTHLY REPORT"
	titleCell.SetStyle(GetTitleStyle())
	titleCell.HMerge = 10

	// Set width for main title row
	for i := 0; i <= 10; i++ {
		sheet.Col(i).Width = titleWidth
	}

	// Add report data
	AddRow(sheet, "Month Of:", report.MonthOf, 120)
	AddRow(sheet, "Worker Name:", report.WorkerName, 120)
	AddRow(sheet, "Area Of Assignment:", report.AreaOfAssignment, 120)
	AddRow(sheet, "Name Of Church:", report.NameOfChurch, 120)

	// Add weekly attendance
	activityRow := sheet.AddRow()
	weeklyAttendanceRow := activityRow.AddCell()
	weeklyAttendanceRow.Value = "WEEKLY ATTENDANCE"
	weeklyAttendanceRow.SetStyle(GetWeeklyAttendanceHeaderStyle())
	weeklyAttendanceRow.HMerge = 10

	// Add Weekly Attendance headers
	activityRow = sheet.AddRow()
	AddMergedCell(activityRow, "Activities", 1, 1)
	addCellWithStyle(activityRow, "Week 1", true)
	addCellWithStyle(activityRow, "Week 2", true)
	addCellWithStyle(activityRow, "Week 3", true)
	addCellWithStyle(activityRow, "Week 4", true)
	addCellWithStyle(activityRow, "Week 5", true)
	addCellWithStyle(activityRow, "Average", true)

	// Apply style to the entire row
	for _, cell := range activityRow.Cells {
		cell.SetStyle(GetWeeklyAttendanceStyle())
	}

	// Add arrays with averages
	AddActivityRow(sheet, "Worship Service:", report.WorshipService, report.WorshipServiceAvg)
	AddActivityRow(sheet, "Sunday School:", report.SundaySchool, report.SundaySchoolAvg)
	AddActivityRow(sheet, "Prayer Meetings:", report.PrayerMeetings, report.PrayerMeetingsAvg)
	AddActivityRow(sheet, "Bible Studies:", report.BibleStudies, report.BibleStudiesAvg)
	AddActivityRow(sheet, "Mens Fellowships:", report.MensFellowships, report.MensFellowshipsAvg)
	AddActivityRow(sheet, "Womens Fellowships:", report.WomensFellowships, report.WomensFellowshipsAvg)
	AddActivityRow(sheet, "Youth Fellowships:", report.YouthFellowships, report.YouthFellowshipsAvg)
	AddActivityRow(sheet, "Child Fellowships:", report.ChildFellowships, report.ChildFellowshipsAvg)
	AddActivityRow(sheet, "Outreach:", report.Outreach, report.OutreachAvg)
	AddActivityRow(sheet, "Training Or Seminars:", report.TrainingOrSeminars, report.TrainingOrSeminarsAvg)
	AddActivityRow(sheet, "Leadership Conferences:", report.LeadershipConferences, report.LeadershipConferencesAvg)
	AddActivityRow(sheet, "Leadership Training:", report.LeadershipTraining, report.LeadershipTrainingAvg)
	AddActivityRow(sheet, "Others:", report.Others, report.OthersAvg)
	AddActivityRow(sheet, "Family Days:", report.FamilyDays, report.FamilyDaysAvg)
	AddActivityRow(sheet, "Tithes And Offerings:", report.TithesAndOfferings, report.TithesAndOfferingsAvg)

	AddActivityRow(sheet, "Home Visited:", report.HomeVisited, report.HomeVisitedAvg)
	AddActivityRow(sheet, "Bible Study Group Led:", report.BibleStudyOrGroupLed, report.BibleStudyOrGroupLedAvg)
	AddActivityRow(sheet, "Sermon/\nMessage Preached:", report.SermonOrMessagePreached, report.SermonOrMessagePreachedAvg)
	AddActivityRow(sheet, "Person Newly Contacted:", report.PersonNewlyContacted, report.PersonNewlyContactedAvg)
	AddActivityRow(sheet, "Person Followed-Up:", report.PersonFollowedUp, report.PersonFollowedUpAvg)
	AddActivityRow(sheet, "Person Led To Christ:", report.PersonLedToChrist, report.PersonLedToChristAvg)

	// Add Names as a comma-separated list
	if len(report.Names) > 0 {
		AddRow(sheet, "Names:", strings.Join(report.Names, ", "), 120)
	} else {
		AddRow(sheet, "Names:", "", 120) // Handle case where Names array is empty
	}

	// Add narrative report
	AddRow(sheet, "Narrative Report:", report.NarrativeReport, 120)
	AddRow(sheet, "Challenges/\nProblems encountered:", report.ChallengesAndProblemEncountered, 120)
	AddRow(sheet, "Prayer Requests:", report.PrayerRequest, 120)

}

func AddRow(sheet *xlsx.Sheet, field, value string, maxCharactersPerCell int) {
	row := sheet.AddRow()
	cellField := row.AddCell()
	cellField.Value = field

	cellValue := row.AddCell()
	if len(value) > maxCharactersPerCell {
		// Insert line breaks in the cell value based on maxCharactersPerCell
		var wrappedValue string
		for i, runeValue := range value {
			wrappedValue += string(runeValue)
			if (i+1)%maxCharactersPerCell == 0 {
				wrappedValue += "\n"
			}
		}
		cellValue.Value = wrappedValue
	} else {
		cellValue.Value = value
	}
}

func GetWrapTextStyle() *xlsx.Style {
	style := xlsx.NewStyle()
	style.Font.Bold = true
	style.Font.Color = "000000" // Black text color
	style.Font.Size = 12
	style.ApplyFont = true
	style.Alignment.WrapText = true
	return style
}

func GetBoldTextStyle() *xlsx.Style {
	style := xlsx.NewStyle()
	style.Font.Bold = true
	style.Font.Color = "000000" // Black text color
	style.Font.Size = 12
	style.ApplyFont = true
	return style
}

// Style functions
func GetTitleStyle() *xlsx.Style {
	style := xlsx.NewStyle()
	style.Font.Bold = true
	style.Font.Size = 13
	style.ApplyFont = true
	style.Alignment.Horizontal = "left"
	return style
}

func GetOrgNameStyle() *xlsx.Style {
	style := xlsx.NewStyle()
	style.Font.Bold = true
	style.Font.Size = 15 // Set font size to 16 (adjust as needed)
	style.ApplyFont = true
	style.Alignment.Horizontal = "left"
	style.Font.Color = "FF0000FF"
	return style
}

func GetWeeklyAttendanceHeaderStyle() *xlsx.Style {
	style := xlsx.NewStyle()
	style.Font.Bold = true
	style.Font.Color = "00000000" // Black text color
	style.Font.Size = 12
	style.ApplyFont = true

	// Set the background color to yellow
	style.Fill = *xlsx.NewFill("solid", "FFFFFF00", "FFFFFF00") // Yellow background in hex
	style.ApplyFill = true

	style.Alignment.Horizontal = "left"
	return style
}
func GetWeeklyAttendanceStyle() *xlsx.Style {
	style := xlsx.NewStyle()
	style.Font.Bold = true
	style.Font.Color = "FFFFFFFF" // White text color in hex (Excel uses ARGB format, where A = Alpha, RGB = Red, Green, Blue)
	style.Font.Size = 11
	style.ApplyFont = true

	// Set the background color to black
	style.Fill = *xlsx.NewFill("solid", "FF000000", "FF000000") // Black background in hex
	style.ApplyFill = true

	style.Alignment.Horizontal = "left"
	return style
}

func AddMergedCell(row *xlsx.Row, value string, hspan, vspan int) {
	cell := row.AddCell()
	cell.Value = value
	cell.Merge(hspan, vspan)
	style := xlsx.NewStyle()
	style.Font.Bold = true
	// style.Fill = *xlsx.NewFill("none", "", "") // Transparent background

	cell.SetStyle(style)
}

func AddActivityRow(sheet *xlsx.Sheet, activity string, values []int, average float64) {
	row := sheet.AddRow()

	// Apply bold text style to the activity label cell
	activityCell := row.AddCell()
	activityCell.Value = activity
	activityCell.SetStyle(GetBoldTextStyle())
	// Ensure there are enough cells for all weeks
	for i := 0; i < 5; i++ {
		if i < len(values) {
			row.AddCell().Value = strconv.Itoa(values[i])
		} else {
			row.AddCell().Value = "" // Blank for weeks with no data
		}
	}

	// Add average
	avgCell := row.AddCell()
	avgCell.Value = fmt.Sprintf("%.2f", average)
}

func addCellWithStyle(row *xlsx.Row, value string, bold bool) {
	cell := row.AddCell()
	cell.Value = value
	style := xlsx.NewStyle()
	if bold {
		style.Font.Bold = true
	}
	cell.SetStyle(style)
}
