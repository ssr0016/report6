package service

import (
	"context"
	"fmt"
	"reports/data/request"
	"reports/data/response"
	"reports/model"
	"reports/repository"
	"time"
)

type ReportServiceImpl struct {
	reportRepository repository.ReportRepository
}

func NewReportServiceImpl(reportRepository repository.ReportRepository) ReportService {
	return &ReportServiceImpl{reportRepository: reportRepository}
}

func (r *ReportServiceImpl) Create(ctx context.Context, request *request.ReportCreateRequest) error {
	// result, err := r.reportRepository.ReportTaken(ctx, 0, request.MonthOf, request.WorkerName)
	// if err != nil {
	// 	return err
	// }

	// if result != nil {
	// 	return errors.New("report already taken")
	// }

	loc, err := time.LoadLocation("Asia/Manila")
	if err != nil {
		return err
	}

	now := time.Now().In(loc)

	report := model.Report{
		MonthOf:                         request.MonthOf,
		WorkerName:                      request.WorkerName,
		AreaOfAssignment:                request.AreaOfAssignment,
		NameOfChurch:                    request.NameOfChurch,
		WorshipService:                  request.WorshipService,
		SundaySchool:                    request.SundaySchool,
		PrayerMeetings:                  request.PrayerMeetings,
		BibleStudies:                    request.BibleStudies,
		MensFellowships:                 request.MensFellowships,
		WomensFellowships:               request.WomensFellowships,
		YouthFellowships:                request.YouthFellowships,
		ChildFellowships:                request.ChildFellowships,
		Outreach:                        request.Outreach,
		TrainingOrSeminars:              request.TrainingOrSeminars,
		LeadershipConferences:           request.LeadershipConferences,
		LeadershipTraining:              request.LeadershipTraining,
		Others:                          request.Others,
		FamilyDays:                      request.FamilyDays,
		TithesAndOfferings:              request.TithesAndOfferings,
		AverageAttendance:               request.AverageAttendance,
		HomeVisited:                     request.HomeVisited,
		BibleStudyOrGroupLed:            request.BibleStudyOrGroupLed,
		SermonOrMessagePreached:         request.SermonOrMessagePreached,
		PersonNewlyContacted:            request.PersonNewlyContacted,
		PersonFollowedUp:                request.PersonFollowedUp,
		PersonLedToChrist:               request.PersonLedToChrist,
		Names:                           request.Names,
		NarrativeReport:                 request.NarrativeReport,
		ChallengesAndProblemEncountered: request.ChallengesAndProblemEncountered,
		PrayerRequest:                   request.PrayerRequest,
		CreatedAt:                       now,
		UpdatedAt:                       now,
	}

	// Save the report using the repository
	err = r.reportRepository.Save(ctx, &report)
	if err != nil {
		return fmt.Errorf("failed to save report: %w", err)
	}

	return nil
}

func (r *ReportServiceImpl) Delete(ctx context.Context, reportId int) error {
	// Retrieve the report by its ID
	report, err := r.reportRepository.FindById(ctx, reportId)
	if err != nil {
		return err // Return error if FindById fails
	}

	// Delete the report using its ID
	err = r.reportRepository.Delete(ctx, report.Id)
	if err != nil {
		return err // Return error if Delete fails
	}

	return nil
}

func (r *ReportServiceImpl) FindAll(ctx context.Context) ([]response.ReportResponse, error) {
	reports, err := r.reportRepository.FindAll(ctx)
	if err != nil {
		return nil, err // Return error if FindAll fails
	}

	var reportResp []response.ReportResponse

	for _, value := range reports {
		report := response.ReportResponse{
			Id:                              value.Id,
			MonthOf:                         value.MonthOf,
			WorkerName:                      value.WorkerName,
			AreaOfAssignment:                value.AreaOfAssignment,
			NameOfChurch:                    value.NameOfChurch,
			WorshipService:                  value.WorshipService,
			SundaySchool:                    value.SundaySchool,
			PrayerMeetings:                  value.PrayerMeetings,
			BibleStudies:                    value.BibleStudies,
			MensFellowships:                 value.MensFellowships,
			WomensFellowships:               value.WomensFellowships,
			YouthFellowships:                value.YouthFellowships,
			ChildFellowships:                value.ChildFellowships,
			Outreach:                        value.Outreach,
			TrainingOrSeminars:              value.TrainingOrSeminars,
			LeadershipConferences:           value.LeadershipConferences,
			LeadershipTraining:              value.LeadershipTraining,
			Others:                          value.Others,
			FamilyDays:                      value.FamilyDays,
			TithesAndOfferings:              value.TithesAndOfferings,
			HomeVisited:                     value.HomeVisited,
			BibleStudyOrGroupLed:            value.BibleStudyOrGroupLed,
			SermonOrMessagePreached:         value.SermonOrMessagePreached,
			PersonNewlyContacted:            value.PersonNewlyContacted,
			PersonFollowedUp:                value.PersonFollowedUp,
			PersonLedToChrist:               value.PersonLedToChrist,
			Names:                           value.Names,
			NarrativeReport:                 value.NarrativeReport,
			ChallengesAndProblemEncountered: value.ChallengesAndProblemEncountered,
			PrayerRequest:                   value.PrayerRequest,
			CreatedAt:                       value.CreatedAt,
			UpdatedAt:                       value.UpdatedAt,
		}

		// Calculate average attendance for each type
		report.WorshipServiceAvg = model.CalculateAverage(value.WorshipService)
		report.SundaySchoolAvg = model.CalculateAverage(value.SundaySchool)
		report.PrayerMeetingsAvg = model.CalculateAverage(value.PrayerMeetings)
		report.BibleStudiesAvg = model.CalculateAverage(value.BibleStudies)
		report.MensFellowshipsAvg = model.CalculateAverage(value.MensFellowships)
		report.WomensFellowshipsAvg = model.CalculateAverage(value.WomensFellowships)
		report.YouthFellowshipsAvg = model.CalculateAverage(value.YouthFellowships)
		report.ChildFellowshipsAvg = model.CalculateAverage(value.ChildFellowships)
		report.OutreachAvg = model.CalculateAverage(value.Outreach)
		report.TrainingOrSeminarsAvg = model.CalculateAverage(value.TrainingOrSeminars)
		report.LeadershipConferencesAvg = model.CalculateAverage(value.LeadershipConferences)
		report.LeadershipTrainingAvg = model.CalculateAverage(value.LeadershipTraining)
		report.OthersAvg = model.CalculateAverage(value.Others)
		report.FamilyDaysAvg = model.CalculateAverage(value.FamilyDays)
		report.TithesAndOfferingsAvg = model.CalculateAverage(value.TithesAndOfferings)
		report.HomeVisitedAvg = model.CalculateAverage(value.HomeVisited)
		report.BibleStudyOrGroupLedAvg = model.CalculateAverage(value.BibleStudyOrGroupLed)
		report.SermonOrMessagePreachedAvg = model.CalculateAverage(value.SermonOrMessagePreached)
		report.PersonNewlyContactedAvg = model.CalculateAverage(value.PersonNewlyContacted)
		report.PersonFollowedUpAvg = model.CalculateAverage(value.PersonFollowedUp)
		report.PersonLedToChristAvg = model.CalculateAverage(value.PersonLedToChrist)

		reportResp = append(reportResp, report)
	}

	return reportResp, nil
}

func (r *ReportServiceImpl) FindById(ctx context.Context, id int) (*response.ReportResponse, error) {
	report, err := r.reportRepository.FindById(ctx, id)
	if err != nil {
		return nil, err // Return error if FindById fails
	}

	reportResp := &response.ReportResponse{
		Id:                              report.Id,
		MonthOf:                         report.MonthOf,
		WorkerName:                      report.WorkerName,
		AreaOfAssignment:                report.AreaOfAssignment,
		NameOfChurch:                    report.NameOfChurch,
		WorshipService:                  report.WorshipService,
		SundaySchool:                    report.SundaySchool,
		PrayerMeetings:                  report.PrayerMeetings,
		BibleStudies:                    report.BibleStudies,
		MensFellowships:                 report.MensFellowships,
		WomensFellowships:               report.WomensFellowships,
		YouthFellowships:                report.YouthFellowships,
		ChildFellowships:                report.ChildFellowships,
		Outreach:                        report.Outreach,
		TrainingOrSeminars:              report.TrainingOrSeminars,
		LeadershipConferences:           report.LeadershipConferences,
		LeadershipTraining:              report.LeadershipTraining,
		Others:                          report.Others,
		FamilyDays:                      report.FamilyDays,
		TithesAndOfferings:              report.TithesAndOfferings,
		HomeVisited:                     report.HomeVisited,
		BibleStudyOrGroupLed:            report.BibleStudyOrGroupLed,
		SermonOrMessagePreached:         report.SermonOrMessagePreached,
		PersonNewlyContacted:            report.PersonNewlyContacted,
		PersonFollowedUp:                report.PersonFollowedUp,
		PersonLedToChrist:               report.PersonLedToChrist,
		Names:                           report.Names,
		NarrativeReport:                 report.NarrativeReport,
		ChallengesAndProblemEncountered: report.ChallengesAndProblemEncountered,
		PrayerRequest:                   report.PrayerRequest,
		CreatedAt:                       report.CreatedAt,
		UpdatedAt:                       report.UpdatedAt,
	}

	// Calculate average attendance for each type
	reportResp.WorshipServiceAvg = model.CalculateAverage(report.WorshipService)
	reportResp.SundaySchoolAvg = model.CalculateAverage(report.SundaySchool)
	reportResp.PrayerMeetingsAvg = model.CalculateAverage(report.PrayerMeetings)
	reportResp.BibleStudiesAvg = model.CalculateAverage(report.BibleStudies)
	reportResp.MensFellowshipsAvg = model.CalculateAverage(report.MensFellowships)
	reportResp.WomensFellowshipsAvg = model.CalculateAverage(report.WomensFellowships)
	reportResp.YouthFellowshipsAvg = model.CalculateAverage(report.YouthFellowships)
	reportResp.ChildFellowshipsAvg = model.CalculateAverage(report.ChildFellowships)
	reportResp.OutreachAvg = model.CalculateAverage(report.Outreach)
	reportResp.TrainingOrSeminarsAvg = model.CalculateAverage(report.TrainingOrSeminars)
	reportResp.LeadershipConferencesAvg = model.CalculateAverage(report.LeadershipConferences)
	reportResp.LeadershipTrainingAvg = model.CalculateAverage(report.LeadershipTraining)
	reportResp.OthersAvg = model.CalculateAverage(report.Others)
	reportResp.FamilyDaysAvg = model.CalculateAverage(report.FamilyDays)
	reportResp.TithesAndOfferingsAvg = model.CalculateAverage(report.TithesAndOfferings)
	reportResp.HomeVisitedAvg = model.CalculateAverage(report.HomeVisited)
	reportResp.BibleStudyOrGroupLedAvg = model.CalculateAverage(report.BibleStudyOrGroupLed)
	reportResp.SermonOrMessagePreachedAvg = model.CalculateAverage(report.SermonOrMessagePreached)
	reportResp.PersonNewlyContactedAvg = model.CalculateAverage(report.PersonNewlyContacted)
	reportResp.PersonFollowedUpAvg = model.CalculateAverage(report.PersonFollowedUp)
	reportResp.PersonLedToChristAvg = model.CalculateAverage(report.PersonLedToChrist)

	return reportResp, nil
}
func (r *ReportServiceImpl) Update(ctx context.Context, request *request.ReportUpdateRequest) error {
	// result, err := r.reportRepository.ReportTaken(ctx, request.Id, request.MonthOf, request.WorkerName)
	// if err != nil {
	// 	return err
	// }

	// // Check if the report exists

	// if len(result) == 0 {
	// 	return errors.New("report already taken")
	// }

	// // Check if there is any report with the same month and worker name that is not the current report
	// if len(result) > 1 || (len(result) == 1 && result[0].Id != request.Id) {
	// 	return errors.New("report already taken")
	// }

	// Retrieve the existing report by ID
	existingReport, err := r.reportRepository.FindById(ctx, request.Id)
	if err != nil {
		return err
	}

	// Update the fields of the existing report entity with request data
	existingReport.MonthOf = request.MonthOf
	existingReport.WorkerName = request.WorkerName
	existingReport.AreaOfAssignment = request.AreaOfAssignment
	existingReport.NameOfChurch = request.NameOfChurch
	existingReport.WorshipService = request.WorshipService
	existingReport.SundaySchool = request.SundaySchool
	existingReport.PrayerMeetings = request.PrayerMeetings
	existingReport.BibleStudies = request.BibleStudies
	existingReport.MensFellowships = request.MensFellowships
	existingReport.WomensFellowships = request.WomensFellowships
	existingReport.YouthFellowships = request.YouthFellowships
	existingReport.ChildFellowships = request.ChildFellowships
	existingReport.Outreach = request.Outreach
	existingReport.TrainingOrSeminars = request.TrainingOrSeminars
	existingReport.LeadershipConferences = request.LeadershipConferences
	existingReport.LeadershipTraining = request.LeadershipTraining
	existingReport.Others = request.Others
	existingReport.FamilyDays = request.FamilyDays
	existingReport.TithesAndOfferings = request.TithesAndOfferings
	existingReport.HomeVisited = request.HomeVisited
	existingReport.BibleStudyOrGroupLed = request.BibleStudyOrGroupLed
	existingReport.SermonOrMessagePreached = request.SermonOrMessagePreached
	existingReport.PersonNewlyContacted = request.PersonNewlyContacted
	existingReport.PersonFollowedUp = request.PersonFollowedUp
	existingReport.PersonLedToChrist = request.PersonLedToChrist
	existingReport.Names = request.Names
	existingReport.NarrativeReport = request.NarrativeReport
	existingReport.ChallengesAndProblemEncountered = request.ChallengesAndProblemEncountered
	existingReport.PrayerRequest = request.PrayerRequest

	err = r.reportRepository.Update(ctx, existingReport)
	if err != nil {
		return err
	}

	return nil
}
