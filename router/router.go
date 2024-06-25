package router

import (
	"net/http"
	"reports/controller"

	"github.com/gin-gonic/gin"
)

func NewRouter(reportController *controller.ReportController) *gin.Engine {
	service := gin.Default()

	service.GET("/", func(ctx *gin.Context) {
		ctx.JSON(http.StatusOK, "Welcome Home!")
	})

	// Api Group
	router := service.Group("/api")

	router.GET("", reportController.FindAll)
	router.POST("", reportController.Create)
	router.GET("/:reportId", reportController.FindById)
	router.PUT("/:reportId", reportController.Update)
	router.DELETE("/:reportId", reportController.Delete)

	return service
}
