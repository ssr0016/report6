package repository

import (
	"context"
	"database/sql"
	"encoding/json"
	"reports/helper"
	"reports/model"
	"strconv"
	"strings"

	_ "github.com/lib/pq"
)

type ReportRepositoryImpl struct {
	Db *sql.DB
}

func NewReportRepository(Db *sql.DB) ReportRepository {
	return &ReportRepositoryImpl{Db: Db}
}

// Delete implements BookRepository
func (r *ReportRepositoryImpl) Delete(ctx context.Context, reportId int) error {
	tx, err := r.Db.BeginTx(ctx, nil)
	if err != nil {
		return err
	}
	defer helper.CommitOrRollback(tx)

	rawSQL := `
		DELETE FROM reports
		WHERE id = $1
	`

	_, err = tx.ExecContext(ctx, rawSQL, reportId)
	if err != nil {
		return err
	}

	return nil
}

// FindAll implements BookRepository
func (r *ReportRepositoryImpl) FindAll(ctx context.Context, query *model.SearchReportQuery) (*model.SearchReportResult, error) {
	tx, err := r.Db.BeginTx(ctx, nil)
	if err != nil {
		return nil, err
	}
	defer helper.CommitOrRollback(tx)

	var rawSQL strings.Builder
	rawSQL.WriteString(`
		SELECT
			id,
			month_of,
			worker_name,
			area_of_assignment,
			name_of_church,
			created_at,
			updated_at,
			worship_service,
			sunday_school,
			prayer_meetings,
			bible_studies,
			mens_fellowships,
			womens_fellowships,
			youth_fellowships,
			child_fellowships,
			outreach,
			training_or_seminars,
			leadership_conferences,
			leadership_training,
			others,
			family_days,
			tithes_and_offerings,
			home_visited,
			bible_study_or_group_led,
			sermon_or_message_preached,
			person_newly_contacted,
			person_followed_up,
			person_led_to_christ,
			names,
			narrative_report,
			challenges_and_problem_encountered,
			prayer_request
		FROM reports t
	`)

	var whereConditions []string
	var whereParams []interface{}
	index := 1

	// Adding dynamic conditions based on query parameters
	if query.MonthOf != "" {
		monthOfParam := "%" + strings.ToLower(query.MonthOf) + "%"
		whereConditions = append(whereConditions, "LOWER(t.month_of) LIKE $"+strconv.Itoa(index))
		whereParams = append(whereParams, monthOfParam)
		index++
	}
	if query.WorkerName != "" {
		workerNameParam := "%" + strings.ToLower(query.WorkerName) + "%"
		whereConditions = append(whereConditions, "LOWER(t.worker_name) LIKE $"+strconv.Itoa(index))
		whereParams = append(whereParams, workerNameParam)
		index++
	}

	if len(whereConditions) > 0 {
		rawSQL.WriteString(" WHERE ")
		rawSQL.WriteString(strings.Join(whereConditions, " AND "))
	}

	rawSQL.WriteString(" ORDER BY t.id") // Replace with your desired ordering column

	// Pagination
	rawSQL.WriteString(" LIMIT $")
	rawSQL.WriteString(strconv.Itoa(index))
	rawSQL.WriteString(" OFFSET $")
	rawSQL.WriteString(strconv.Itoa(index + 1))

	// Append pagination parameters to args slice
	whereParams = append(whereParams, query.PerPage, (query.Page-1)*query.PerPage)

	// Execute query
	rows, err := tx.QueryContext(ctx, rawSQL.String(), whereParams...)
	if err != nil {
		return nil, err
	}
	defer rows.Close()

	var reports []*model.Report // Changed to []*model.Report to match SearchReportResult.Reports

	for rows.Next() {
		var report model.Report
		var (
			worshipServiceJSON        []byte
			sundaySchoolJSON          []byte
			prayerMeetingsJSON        []byte
			bibleStudiesJSON          []byte
			mensFellowshipsJSON       []byte
			womensFellowshipsJSON     []byte
			youthFellowshipsJSON      []byte
			childFellowshipsJSON      []byte
			outreachJSON              []byte
			trainingOrSeminarsJSON    []byte
			leadershipConferencesJSON []byte
			leadershipTrainingJSON    []byte
			othersJSON                []byte
			familyDaysJSON            []byte
			tithesAndOfferingsJSON    []byte
			homeVisitedJSON           []byte
			bibleStudyOrGroupLedJSON  []byte
			sermonOrMessageJSON       []byte
			personNewlyContactedJSON  []byte
			personFollowedUpJSON      []byte
			personLedToChristJSON     []byte
			namesJSON                 []byte
		)

		// Scan row into variables
		err := rows.Scan(
			&report.Id,
			&report.MonthOf,
			&report.WorkerName,
			&report.AreaOfAssignment,
			&report.NameOfChurch,
			&report.CreatedAt,
			&report.UpdatedAt,
			&worshipServiceJSON,
			&sundaySchoolJSON,
			&prayerMeetingsJSON,
			&bibleStudiesJSON,
			&mensFellowshipsJSON,
			&womensFellowshipsJSON,
			&youthFellowshipsJSON,
			&childFellowshipsJSON,
			&outreachJSON,
			&trainingOrSeminarsJSON,
			&leadershipConferencesJSON,
			&leadershipTrainingJSON,
			&othersJSON,
			&familyDaysJSON,
			&tithesAndOfferingsJSON,
			&homeVisitedJSON,
			&bibleStudyOrGroupLedJSON,
			&sermonOrMessageJSON,
			&personNewlyContactedJSON,
			&personFollowedUpJSON,
			&personLedToChristJSON,
			&namesJSON,
			&report.NarrativeReport,
			&report.ChallengesAndProblemEncountered,
			&report.PrayerRequest,
		)
		if err != nil {
			return nil, err
		}

		// Unmarshal JSONB fields into respective slices or fields
		if worshipServiceJSON != nil {
			if err := json.Unmarshal(worshipServiceJSON, &report.WorshipService); err != nil {
				return nil, err
			}
		}
		// Repeat for other JSONB fields...

		reports = append(reports, &report) // Append pointer to report
	}

	// Check for any error during iteration
	if err := rows.Err(); err != nil {
		return nil, err
	}

	// Construct the result object
	result := &model.SearchReportResult{
		TotalCount: len(reports), // Assuming you want total count of items fetched
		Reports:    reports,
		Page:       query.Page,
		PerPage:    query.PerPage,
	}

	return result, nil
}

// FindById implements BookRepository
func (r *ReportRepositoryImpl) FindById(ctx context.Context, id int) (*model.Report, error) {
	tx, err := r.Db.BeginTx(ctx, nil)
	if err != nil {
		return nil, err
	}
	defer helper.CommitOrRollback(tx)

	// Construct the dynamic SELECT statement based on non-zero fields
	rawSQL := `
		SELECT 
			id,
			month_of,
			worker_name,
			area_of_assignment,
			name_of_church,
			created_at,
			updated_at,
			worship_service,
			sunday_school,
			prayer_meetings,
			bible_studies,
			mens_fellowships,
			womens_fellowships,
			youth_fellowships,
			child_fellowships,
			outreach,
			training_or_seminars,
			leadership_conferences,
			leadership_training,
			others,
			family_days,
			tithes_and_offerings,
			home_visited,
			bible_study_or_group_led,
			sermon_or_message_preached,
			person_newly_contacted,
			person_followed_up,
			person_led_to_christ,
			names,
			narrative_report,
			challenges_and_problem_encountered,
			prayer_request
		FROM reports
		WHERE id = $1
	`

	var report model.Report
	var (
		worshipServiceJSON        []byte
		sundaySchoolJSON          []byte
		prayerMeetingsJSON        []byte
		bibleStudiesJSON          []byte
		mensFellowshipsJSON       []byte
		womensFellowshipsJSON     []byte
		youthFellowshipsJSON      []byte
		childFellowshipsJSON      []byte
		outreachJSON              []byte
		trainingOrSeminarsJSON    []byte
		leadershipConferencesJSON []byte
		leadershipTrainingJSON    []byte
		othersJSON                []byte
		familyDaysJSON            []byte
		tithesAndOfferingsJSON    []byte
		homeVisitedJSON           []byte
		bibleStudyOrGroupLedJSON  []byte
		sermonOrMessageJSON       []byte
		personNewlyContactedJSON  []byte
		personFollowedUpJSON      []byte
		personLedToChristJSON     []byte
		namesJSON                 []byte
	)

	// Execute the query
	err = tx.QueryRowContext(ctx, rawSQL, id).Scan(
		&report.Id,
		&report.MonthOf,
		&report.WorkerName,
		&report.AreaOfAssignment,
		&report.NameOfChurch,
		&report.CreatedAt,
		&report.UpdatedAt,
		&worshipServiceJSON,
		&sundaySchoolJSON,
		&prayerMeetingsJSON,
		&bibleStudiesJSON,
		&mensFellowshipsJSON,
		&womensFellowshipsJSON,
		&youthFellowshipsJSON,
		&childFellowshipsJSON,
		&outreachJSON,
		&trainingOrSeminarsJSON,
		&leadershipConferencesJSON,
		&leadershipTrainingJSON,
		&othersJSON,
		&familyDaysJSON,
		&tithesAndOfferingsJSON,
		&homeVisitedJSON,
		&bibleStudyOrGroupLedJSON,
		&sermonOrMessageJSON,
		&personNewlyContactedJSON,
		&personFollowedUpJSON,
		&personLedToChristJSON,
		&namesJSON,
		&report.NarrativeReport,
		&report.ChallengesAndProblemEncountered,
		&report.PrayerRequest,
	)
	if err != nil {
		return nil, err
	}

	// Unmarshal JSONB fields into their respective slices
	if worshipServiceJSON != nil {
		if err := json.Unmarshal(worshipServiceJSON, &report.WorshipService); err != nil {
			return nil, err
		}
	}

	if sundaySchoolJSON != nil {
		if err := json.Unmarshal(sundaySchoolJSON, &report.SundaySchool); err != nil {
			return nil, err
		}
	}

	if prayerMeetingsJSON != nil {
		if err := json.Unmarshal(prayerMeetingsJSON, &report.PrayerMeetings); err != nil {
			return nil, err
		}
	}

	if bibleStudiesJSON != nil {
		if err := json.Unmarshal(bibleStudiesJSON, &report.BibleStudies); err != nil {
			return nil, err
		}
	}

	if mensFellowshipsJSON != nil {
		if err := json.Unmarshal(mensFellowshipsJSON, &report.MensFellowships); err != nil {
			return nil, err
		}
	}

	if womensFellowshipsJSON != nil {
		if err := json.Unmarshal(womensFellowshipsJSON, &report.WomensFellowships); err != nil {
			return nil, err
		}
	}

	if youthFellowshipsJSON != nil {
		if err := json.Unmarshal(youthFellowshipsJSON, &report.YouthFellowships); err != nil {
			return nil, err
		}
	}

	if childFellowshipsJSON != nil {
		if err := json.Unmarshal(childFellowshipsJSON, &report.ChildFellowships); err != nil {
			return nil, err
		}
	}

	if outreachJSON != nil {
		if err := json.Unmarshal(outreachJSON, &report.Outreach); err != nil {
			return nil, err
		}
	}

	if trainingOrSeminarsJSON != nil {
		if err := json.Unmarshal(trainingOrSeminarsJSON, &report.TrainingOrSeminars); err != nil {
			return nil, err
		}
	}

	if leadershipConferencesJSON != nil {
		if err := json.Unmarshal(leadershipConferencesJSON, &report.LeadershipConferences); err != nil {
			return nil, err
		}
	}

	if leadershipTrainingJSON != nil {
		if err := json.Unmarshal(leadershipTrainingJSON, &report.LeadershipTraining); err != nil {
			return nil, err
		}
	}

	if othersJSON != nil {
		if err := json.Unmarshal(othersJSON, &report.Others); err != nil {
			return nil, err
		}
	}

	if familyDaysJSON != nil {
		if err := json.Unmarshal(familyDaysJSON, &report.FamilyDays); err != nil {
			return nil, err
		}
	}

	if tithesAndOfferingsJSON != nil {
		if err := json.Unmarshal(tithesAndOfferingsJSON, &report.TithesAndOfferings); err != nil {
			return nil, err
		}
	}

	if homeVisitedJSON != nil {
		if err := json.Unmarshal(homeVisitedJSON, &report.HomeVisited); err != nil {
			return nil, err
		}
	}

	if bibleStudyOrGroupLedJSON != nil {
		if err := json.Unmarshal(bibleStudyOrGroupLedJSON, &report.BibleStudyOrGroupLed); err != nil {
			return nil, err
		}
	}

	if sermonOrMessageJSON != nil {
		if err := json.Unmarshal(sermonOrMessageJSON, &report.SermonOrMessagePreached); err != nil {
			return nil, err
		}
	}

	if personNewlyContactedJSON != nil {
		if err := json.Unmarshal(personNewlyContactedJSON, &report.PersonNewlyContacted); err != nil {
			return nil, err
		}
	}

	if personFollowedUpJSON != nil {
		if err := json.Unmarshal(personFollowedUpJSON, &report.PersonFollowedUp); err != nil {
			return nil, err
		}
	}

	if personLedToChristJSON != nil {
		if err := json.Unmarshal(personLedToChristJSON, &report.PersonLedToChrist); err != nil {
			return nil, err
		}
	}

	if namesJSON != nil {
		if err := json.Unmarshal(namesJSON, &report.Names); err != nil {
			return nil, err
		}
	}

	return &report, nil
}

// Save implements BookRepository
func (r *ReportRepositoryImpl) Save(ctx context.Context, report *model.Report) error {
	tx, err := r.Db.Begin()
	if err != nil {
		return err
	}
	defer helper.CommitOrRollback(tx)

	// Marshal arrays to JSONB
	worshipServiceJSON, err := json.Marshal(report.WorshipService)
	if err != nil {
		return err
	}

	sundaySchoolJSON, err := json.Marshal(report.SundaySchool)
	if err != nil {
		return err
	}

	prayerMeetingsJSON, err := json.Marshal(report.PrayerMeetings)
	if err != nil {
		return err
	}

	bibleStudiesJSON, err := json.Marshal(report.BibleStudies)
	if err != nil {
		return err
	}

	mensFellowshipsJSON, err := json.Marshal(report.MensFellowships)
	if err != nil {
		return err
	}

	womensFellowshipsJSON, err := json.Marshal(report.WomensFellowships)
	if err != nil {
		return err
	}

	youthFellowshipsJSON, err := json.Marshal(report.YouthFellowships)
	if err != nil {
		return err
	}

	childFellowshipsJSON, err := json.Marshal(report.ChildFellowships)
	if err != nil {
		return err
	}

	outreachJSON, err := json.Marshal(report.Outreach)
	if err != nil {
		return err
	}

	trainingOrSeminarsJSON, err := json.Marshal(report.TrainingOrSeminars)
	if err != nil {
		return err
	}

	leadershipConferencesJSON, err := json.Marshal(report.LeadershipConferences)
	if err != nil {
		return err
	}

	leadershipTrainingJSON, err := json.Marshal(report.LeadershipTraining)
	if err != nil {
		return err
	}

	othersJSON, err := json.Marshal(report.Others)
	if err != nil {
		return err
	}

	familyDaysJSON, err := json.Marshal(report.FamilyDays)
	if err != nil {
		return err
	}

	tithesAndOfferingsJSON, err := json.Marshal(report.TithesAndOfferings)
	if err != nil {
		return err
	}

	homeVisitedJSON, err := json.Marshal(report.HomeVisited)
	if err != nil {
		return err
	}

	bibleStudyOrGroupLedJSON, err := json.Marshal(report.BibleStudyOrGroupLed)
	if err != nil {
		return err
	}

	sermonOrMessagePreachedJSON, err := json.Marshal(report.SermonOrMessagePreached)
	if err != nil {
		return err
	}

	personNewlyContactedJSON, err := json.Marshal(report.PersonNewlyContacted)
	if err != nil {
		return err
	}

	personFollowedUpJSON, err := json.Marshal(report.PersonFollowedUp)
	if err != nil {
		return err
	}

	personLedToChristJSON, err := json.Marshal(report.PersonLedToChrist)
	if err != nil {
		return err
	}

	namesJSON, err := json.Marshal(report.Names)
	if err != nil {
		return err
	}

	rawSQL := `
		INSERT INTO reports (
			month_of,
			worker_name,
			area_of_assignment,
			name_of_church,
			worship_service,
			sunday_school,
			prayer_meetings,
			bible_studies,
			mens_fellowships,
			womens_fellowships,
			youth_fellowships,
			child_fellowships,
			outreach,
			training_or_seminars,
			leadership_conferences,
			leadership_training,
			others,
			family_days,
			tithes_and_offerings,
			home_visited,
			bible_study_or_group_led,
			sermon_or_message_preached,
			person_newly_contacted,
			person_followed_up,
			person_led_to_christ,
			names,
			narrative_report,
			challenges_and_problem_encountered,
			prayer_request,
			created_at,
			updated_at
		) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, $18, $19, $20, $21, $22, $23, $24, $25, $26, $27, $28, $29, $30, $31)
	`

	_, err = tx.ExecContext(ctx, rawSQL,
		report.MonthOf,
		report.WorkerName,
		report.AreaOfAssignment,
		report.NameOfChurch,
		worshipServiceJSON,
		sundaySchoolJSON,
		prayerMeetingsJSON,
		bibleStudiesJSON,
		mensFellowshipsJSON,
		womensFellowshipsJSON,
		youthFellowshipsJSON,
		childFellowshipsJSON,
		outreachJSON,
		trainingOrSeminarsJSON,
		leadershipConferencesJSON,
		leadershipTrainingJSON,
		othersJSON,
		familyDaysJSON,
		tithesAndOfferingsJSON,
		homeVisitedJSON,
		bibleStudyOrGroupLedJSON,
		sermonOrMessagePreachedJSON,
		personNewlyContactedJSON,
		personFollowedUpJSON,
		personLedToChristJSON,
		namesJSON,
		report.NarrativeReport,
		report.ChallengesAndProblemEncountered,
		report.PrayerRequest,
		report.CreatedAt,
		report.UpdatedAt,
	)
	if err != nil {
		return err
	}

	return nil
}

// Update implements BookRepository
func (r *ReportRepositoryImpl) Update(ctx context.Context, report *model.Report) error {
	tx, err := r.Db.Begin()
	if err != nil {
		return err
	}
	defer helper.CommitOrRollback(tx)

	rawSQL := `
        UPDATE reports SET
            month_of = $1,
            worker_name = $2,
            area_of_assignment = $3,
            name_of_church = $4,
            worship_service = $5,
            sunday_school = $6,
            prayer_meetings = $7,
            bible_studies = $8,
            mens_fellowships = $9,
            womens_fellowships = $10,
            youth_fellowships = $11,
            child_fellowships = $12,
            outreach = $13,
            training_or_seminars = $14,
            leadership_conferences = $15,
            leadership_training = $16,
            others = $17,
            family_days = $18,
            tithes_and_offerings = $19,
            home_visited = $20,
            bible_study_or_group_led = $21,
            sermon_or_message_preached = $22,
            person_newly_contacted = $23,
            person_followed_up = $24,
            person_led_to_christ = $25,
            names = $26,
            narrative_report = $27,
            challenges_and_problem_encountered = $28,
            prayer_request = $29,
            updated_at = $30
        WHERE 
            id = $32
    `

	// Marshal arrays to JSON
	worshipServiceJSON, err := json.Marshal(report.WorshipService)
	if err != nil {
		return err
	}
	sundaySchoolJSON, err := json.Marshal(report.SundaySchool)
	if err != nil {
		return err
	}
	prayerMeetingsJSON, err := json.Marshal(report.PrayerMeetings)
	if err != nil {
		return err
	}
	bibleStudiesJSON, err := json.Marshal(report.BibleStudies)
	if err != nil {
		return err
	}
	mensFellowshipsJSON, err := json.Marshal(report.MensFellowships)
	if err != nil {
		return err
	}
	womensFellowshipsJSON, err := json.Marshal(report.WomensFellowships)
	if err != nil {
		return err
	}
	youthFellowshipsJSON, err := json.Marshal(report.YouthFellowships)
	if err != nil {
		return err
	}
	childFellowshipsJSON, err := json.Marshal(report.ChildFellowships)
	if err != nil {
		return err
	}
	outreachJSON, err := json.Marshal(report.Outreach)
	if err != nil {
		return err
	}
	trainingOrSeminarsJSON, err := json.Marshal(report.TrainingOrSeminars)
	if err != nil {
		return err
	}
	leadershipConferencesJSON, err := json.Marshal(report.LeadershipConferences)
	if err != nil {
		return err
	}
	leadershipTrainingJSON, err := json.Marshal(report.LeadershipTraining)
	if err != nil {
		return err
	}
	othersJSON, err := json.Marshal(report.Others)
	if err != nil {
		return err
	}
	familyDaysJSON, err := json.Marshal(report.FamilyDays)
	if err != nil {
		return err
	}
	tithesAndOfferingsJSON, err := json.Marshal(report.TithesAndOfferings)
	if err != nil {
		return err
	}
	homeVisitedJSON, err := json.Marshal(report.HomeVisited)
	if err != nil {
		return err
	}
	bibleStudyOrGroupLedJSON, err := json.Marshal(report.BibleStudyOrGroupLed)
	if err != nil {
		return err
	}
	sermonOrMessagePreachedJSON, err := json.Marshal(report.SermonOrMessagePreached)
	if err != nil {
		return err
	}
	personNewlyContactedJSON, err := json.Marshal(report.PersonNewlyContacted)
	if err != nil {
		return err
	}
	personFollowedUpJSON, err := json.Marshal(report.PersonFollowedUp)
	if err != nil {
		return err
	}
	personLedToChristJSON, err := json.Marshal(report.PersonLedToChrist)
	if err != nil {
		return err
	}
	namesJSON, err := json.Marshal(report.Names)
	if err != nil {
		return err
	}

	// Execute the update query
	_, err = tx.ExecContext(ctx, rawSQL,
		report.MonthOf,
		report.WorkerName,
		report.AreaOfAssignment,
		report.NameOfChurch,
		worshipServiceJSON,
		sundaySchoolJSON,
		prayerMeetingsJSON,
		bibleStudiesJSON,
		mensFellowshipsJSON,
		womensFellowshipsJSON,
		youthFellowshipsJSON,
		childFellowshipsJSON,
		outreachJSON,
		trainingOrSeminarsJSON,
		leadershipConferencesJSON,
		leadershipTrainingJSON,
		othersJSON,
		familyDaysJSON,
		tithesAndOfferingsJSON,
		homeVisitedJSON,
		bibleStudyOrGroupLedJSON,
		sermonOrMessagePreachedJSON,
		personNewlyContactedJSON,
		personFollowedUpJSON,
		personLedToChristJSON,
		namesJSON,
		report.NarrativeReport,
		report.ChallengesAndProblemEncountered,
		report.PrayerRequest,
		report.UpdatedAt,
		report.Id,
	)
	if err != nil {
		return err
	}

	return nil
}

func (r *ReportRepositoryImpl) ReportTaken(ctx context.Context, id int, monthOf, workerName string) ([]*model.Report, error) {
	var reports []*model.Report

	rawSQL := `
		SELECT
			id, 
			month_of,
			worker_name
		FROM reports
		WHERE month_of = $1 OR
		worker_name = $2
	`

	rows, err := r.Db.QueryContext(ctx, rawSQL, monthOf, workerName)
	if err != nil {
		return nil, err
	}
	defer rows.Close()

	for rows.Next() {
		var report model.Report
		if err := rows.Scan(
			&report.Id,
			&report.MonthOf,
			&report.WorkerName,
		); err != nil {
			return nil, err
		}
		reports = append(reports, &report)
	}

	return reports, nil
}
