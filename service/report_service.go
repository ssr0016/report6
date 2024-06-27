package service

import (
	"context"
	"reports/data/request"
	"reports/data/response"
	"reports/model"
)

type ReportService interface {
	Create(ctx context.Context, request *request.ReportCreateRequest) error
	Update(ctx context.Context, request *request.ReportUpdateRequest) error
	Delete(ctx context.Context, reportId int) error
	FindById(ctx context.Context, reportId int) (*response.ReportResponse, error)
	FindAll(ctx context.Context, query *model.SearchReportQuery) (*model.SearchReportResult, error)
}
