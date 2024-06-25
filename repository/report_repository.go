package repository

import (
	"context"
	"reports/model"
)

type ReportRepository interface {
	Save(ctx context.Context, report *model.Report) error
	Update(ctx context.Context, report *model.Report) error
	Delete(ctx context.Context, reportId int) error
	FindById(ctx context.Context, reportId int) (*model.Report, error)
	FindAll(ctx context.Context) ([]model.Report, error)
	ReportTaken(ctx context.Context, id int, monthOf, workerName string) ([]*model.Report, error)
}
