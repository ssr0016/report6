package controller

import (
	"context"
	"net/http"
	"reports/data/request"
	"reports/model"
	"reports/service"
	"reports/utils"

	"strconv"

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
	// Parse query parameters
	query := model.SearchReportQuery{
		MonthOf:    ctx.DefaultQuery("month_of", ""),
		WorkerName: ctx.Query("worker_name"),
		Page:       parsePage(ctx.DefaultQuery("page", "1")),
		PerPage:    parsePerPage(ctx.DefaultQuery("per_page", "10")),
	}

	// Call the service layer to fetch data
	result, err := controller.reportService.FindAll(ctx, &query)
	if err != nil {
		ctx.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to fetch reports", "details": err.Error()})
		return
	}

	// Return the JSON response
	ctx.JSON(http.StatusOK, gin.H{
		"reports": gin.H{
			"total_count": result.TotalCount,
			"reports":     result.Reports,
			"page":        query.Page,
			"per_page":    query.PerPage,
		},
	})
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
	utils.AddReportToSheet(sheet, report)

	// Set the response headers
	ctx.Header("Content-Description", "File Transfer")
	ctx.Header("Content-Disposition", "attachment; filename=report.xlsx")
	ctx.Header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

	// Write the file to the response
	if err := file.Write(ctx.Writer); err != nil {
		ctx.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to write Excel file"})
	}
}

func parsePage(pageStr string) int {
	page, err := strconv.Atoi(pageStr)
	if err != nil || page < 1 {
		return 1 // Default to page 1 if invalid or not provided
	}
	return page
}

// Helper function to parse per_page from query parameter
func parsePerPage(perPageStr string) int {
	perPage, err := strconv.Atoi(perPageStr)
	if err != nil || perPage < 1 {
		return 10 // Default to 10 per page if invalid or not provided
	}
	return perPage
}
