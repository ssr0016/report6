package controller

import (
	"net/http"
	"reports/data/request"
	"reports/service"
	"strconv"

	"github.com/gin-gonic/gin"
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
