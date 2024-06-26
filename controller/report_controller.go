package controller

import (
	"context"
	"fmt"
	"net/http"
	"reports/data/request"
	"reports/data/response"
	"reports/service"
	"strconv"
	"strings"

	"github.com/gin-gonic/gin"
	"github.com/tealeg/xlsx"
)

type ReportController struct {
	reportService service.ReportService
}

func NewReportController(reportService service.ReportService) *ReportController {
	return &ReportController{reportService: reportService}
}

func (controller *ReportController) Create(ctx *gin.Context) {
	var req request.ReportCreateRequest
	if err := ctx.ShouldBindJSON(&req); err != nil {
		ctx.JSON(http.StatusBadRequest, gin.H{"error": "Invalid request body"})
		return
	}

	if err := req.Validate(); err != nil {
		ctx.JSON(http.StatusBadRequest, gin.H{"error": err.Error()})
		return
	}

	if err := controller.reportService.Create(ctx.Request.Context(), &req); err != nil {
		ctx.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to create report", "details": err.Error()})
		return
	}

	ctx.JSON(http.StatusOK, gin.H{"message": "Report created successfully"})
}

func (controller *ReportController) FindById(ctx *gin.Context) {
	reportId, err := strconv.Atoi(ctx.Param("reportId"))
	if err != nil {
		ctx.JSON(http.StatusBadRequest, gin.H{"error": "Invalid report ID"})
		return
	}

	report, err := controller.reportService.FindById(ctx, reportId)
	if err != nil {
		ctx.JSON(http.StatusNotFound, gin.H{"error": "Report not found", "details": err.Error()})
		return
	}

	ctx.JSON(http.StatusOK, gin.H{"report": report})
}

func (controller *ReportController) FindAll(ctx *gin.Context) {
	reports, err := controller.reportService.FindAll(ctx)
	if err != nil {
		ctx.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to fetch reports", "details": err.Error()})
		return
	}

	ctx.JSON(http.StatusOK, gin.H{"reports": reports})
}

func (controller *ReportController) Delete(ctx *gin.Context) {
	reportId, err := strconv.Atoi(ctx.Param("reportId"))
	if err != nil {
		ctx.JSON(http.StatusBadRequest, gin.H{"error": "Invalid report ID"})
		return
	}

	if err := controller.reportService.Delete(ctx, reportId); err != nil {
		ctx.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to delete report", "details": err.Error()})
		return
	}

	ctx.JSON(http.StatusOK, gin.H{"message": "Report deleted successfully"})
}

func (controller *ReportController) Update(ctx *gin.Context) {
	var req request.ReportUpdateRequest

	if err := ctx.ShouldBindJSON(&req); err != nil {
		ctx.JSON(http.StatusBadRequest, gin.H{"error": "Invalid request body"})
		return
	}

	reportId, err := strconv.Atoi(ctx.Param("reportId"))
	if err != nil {
		ctx.JSON(http.StatusBadRequest, gin.H{"error": "Invalid report ID"})
		return
	}

	req.Id = reportId

	// Validate request parameters
	if err := req.Validate(); err != nil {
		ctx.JSON(http.StatusBadRequest, gin.H{"error": err.Error()})
		return
	}

	// Call service layer to update the report
	if err := controller.reportService.Update(ctx, &req); err != nil {
		ctx.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to update report", "details": err.Error()})
		return
	}

	// Respond with only success message
	ctx.JSON(http.StatusOK, gin.H{"message": "Report updated successfully"})
}

func (controller *ReportController) ExportReport(ctx *gin.Context) {
	reportId, err := strconv.Atoi(ctx.Param("reportId"))
	if err != nil {
		ctx.JSON(http.StatusBadRequest, gin.H{"error": "Invalid report ID"})
		return
	}

	report, err := controller.reportService.FindById(context.Background(), reportId)
	if err != nil {
		ctx.JSON(http.StatusInternalServerError, gin.H{"error": err.Error()})
		return
	}

	// Create a new Excel file
	file := xlsx.NewFile()
	sheet, err := file.AddSheet("Report")
	if err != nil {
		ctx.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to create Excel sheet"})
		return
	}

	// Add data to the Excel sheet
	addReportToSheet(sheet, report)

	// Set the response headers
	ctx.Header("Content-Description", "File Transfer")
	ctx.Header("Content-Disposition", "attachment; filename=report.xlsx")
	ctx.Header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

	// Write the file to the response
	if err := file.Write(ctx.Writer); err != nil {
		ctx.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to write Excel file"})
	}
}

func addReportToSheet(sheet *xlsx.Sheet, report *response.ReportResponse) {
	orgNameRow := sheet.AddRow()
	orgNameCell := orgNameRow.AddCell()
	orgNameCell.Value = "ANG MANANAMPALATAYANG GUMAWA"
	orgNameCell.SetStyle(getOrgNameStyle())

	// Add main title
	titleRow := sheet.AddRow()
	titleCell := titleRow.AddCell()

	titleCell.Value = "NATIONAL WORKERS' MONTHLY REPORT"
	titleCell.SetStyle(getTitleStyle())

	// // Add headers
	// headers := []string{"Field", "Value"}
	// row := sheet.AddRow()
	// for _, header := range headers {
	// 	cell := row.AddCell()
	// 	cell.Value = header
	// }

	addRow(sheet, "ID", strconv.Itoa(report.Id))
	addRow(sheet, "Month Of:", report.MonthOf)
	addRow(sheet, "Worker Name:", report.WorkerName)
	addRow(sheet, "Area Of Assignment:", report.AreaOfAssignment)
	addRow(sheet, "Name Of Church:", report.NameOfChurch)

	// Add Weekly Attendance header (bold)
	weeklyAttendanceRow := sheet.AddRow()
	weeklyAttendanceCell := weeklyAttendanceRow.AddCell()
	weeklyAttendanceCell.Value = "Weekly Attendance"
	weeklyAttendanceCell.SetStyle(getWeeklyAttendaceStyle())

	// Add weekly attendance
	addReportHeaders(sheet)

	// Add arrays with averages
	addIntArrayWithAvg(sheet, "Worship Service:", report.WorshipService, report.WorshipServiceAvg)
	addIntArrayWithAvg(sheet, "Sunday School:", report.SundaySchool, report.SundaySchoolAvg)
	addIntArrayWithAvg(sheet, "Prayer Meetings:", report.PrayerMeetings, report.PrayerMeetingsAvg)
	addIntArrayWithAvg(sheet, "Bible Studies:", report.BibleStudies, report.BibleStudiesAvg)
	addIntArrayWithAvg(sheet, "Mens Fellowships:", report.MensFellowships, report.MensFellowshipsAvg)
	addIntArrayWithAvg(sheet, "Womens Fellowships:", report.WomensFellowships, report.WomensFellowshipsAvg)
	addIntArrayWithAvg(sheet, "Youth Fellowships:", report.YouthFellowships, report.YouthFellowshipsAvg)
	addIntArrayWithAvg(sheet, "Child Fellowships:", report.ChildFellowships, report.ChildFellowshipsAvg)
	addIntArrayWithAvg(sheet, "Outreach:", report.Outreach, report.OutreachAvg)
	addIntArrayWithAvg(sheet, "Training Or Seminars:", report.TrainingOrSeminars, report.TrainingOrSeminarsAvg)
	addIntArrayWithAvg(sheet, "Leadership Conferences:", report.LeadershipConferences, report.LeadershipConferencesAvg)
	addIntArrayWithAvg(sheet, "Leadership Training:", report.LeadershipTraining, report.LeadershipTrainingAvg)
	addIntArrayWithAvg(sheet, "Others:", report.Others, report.OthersAvg)
	addIntArrayWithAvg(sheet, "Family Days:", report.FamilyDays, report.FamilyDaysAvg)
	addIntArrayWithAvg(sheet, "Tithes And Offerings:", report.TithesAndOfferings, report.TithesAndOfferingsAvg)

	addIntArrayWithAvg(sheet, "Home Visited:", report.HomeVisited, report.HomeVisitedAvg)
	addIntArrayWithAvg(sheet, "Bible Study Or Group Led:", report.BibleStudyOrGroupLed, report.BibleStudyOrGroupLedAvg)
	addIntArrayWithAvg(sheet, "Sermon Or Message Preached:", report.SermonOrMessagePreached, report.SermonOrMessagePreachedAvg)
	addIntArrayWithAvg(sheet, "Person Newly Contacted:", report.PersonNewlyContacted, report.PersonNewlyContactedAvg)
	addIntArrayWithAvg(sheet, "Person Followed-Up:", report.PersonFollowedUp, report.PersonFollowedUpAvg)
	addIntArrayWithAvg(sheet, "Person Led To Christ:", report.PersonLedToChrist, report.PersonLedToChristAvg)

	// Add narrative report
	addRow(sheet, "Narrative Report:", report.NarrativeReport)
	addRow(sheet, "Challenges/Problems encountered:", report.ChallengesAndProblemEncountered)
	addRow(sheet, "Prayer Requests:", report.PrayerRequest)

}

func addRow(sheet *xlsx.Sheet, field, value string) {
	row := sheet.AddRow()
	row.AddCell().Value = field
	row.AddCell().Value = value
}

func addIntArrayWithAvg(sheet *xlsx.Sheet, field string, values []int, avg float64) {
	// Add array values
	row := sheet.AddRow()
	row.AddCell().Value = field
	valuesStr := strings.Trim(strings.Replace(fmt.Sprint(values), " ", ", ", -1), "[]")
	row.AddCell().Value = valuesStr

	// Add average in a separate cell
	avgCell := row.AddCell()
	avgCell.Value = fmt.Sprintf("(Average: %.2f)", avg)
}

// Style functions
func getTitleStyle() *xlsx.Style {
	style := xlsx.NewStyle()
	style.Font.Bold = true
	style.ApplyFont = true
	style.Alignment.Horizontal = "center"
	return style
}

func getOrgNameStyle() *xlsx.Style {
	style := xlsx.NewStyle()
	style.Font.Bold = true
	style.Font.Size = 16 // Set font size to 16 (adjust as needed)
	style.ApplyFont = true
	style.Alignment.Horizontal = "center"
	style.Font.Color = "FF0000FF"
	return style
}

func getWeeklyAttendaceStyle() *xlsx.Style {
	style := xlsx.NewStyle()
	style.Font.Bold = true
	return style
}

func addReportHeaders(sheet *xlsx.Sheet) {
	row := sheet.AddRow()
	addMergedCell(row, "Activities", 1, 1)
	addMergedCell(row, "Week 1", 1, 1)
	addMergedCell(row, "Week 2", 1, 1)
	addMergedCell(row, "Week 3", 1, 1)
	addMergedCell(row, "Week 4", 1, 1)
	addMergedCell(row, "Week 5", 1, 1)
	addMergedCell(row, "Average", 1, 1)
}

func addMergedCell(row *xlsx.Row, value string, hspan, vspan int) {
	cell := row.AddCell()
	cell.Value = value
	cell.Merge(hspan, vspan)
	style := xlsx.NewStyle()
	style.Font.Bold = true
	style.Fill = *xlsx.NewFill("none", "", "") // Transparent background
	cell.SetStyle(style)
}
