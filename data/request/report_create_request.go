package request

import "errors"

type ReportCreateRequest struct {
	MonthOf                         string   `json:"month_of" validate:"required"`
	WorkerName                      string   `json:"worker_name" validate:"required"`
	AreaOfAssignment                string   `json:"area_of_assignment" validate:"required"`
	NameOfChurch                    string   `json:"name_of_church" validate:"required"`
	WorshipService                  []int    `json:"worship_service" validate:"required"`
	SundaySchool                    []int    `json:"sunday_school" validate:"required"`
	PrayerMeetings                  []int    `json:"prayer_meetings,omitempty"`
	BibleStudies                    []int    `json:"bible_studies,omitempty"`
	MensFellowships                 []int    `json:"mens_fellowships,omitempty"`
	WomensFellowships               []int    `json:"womens_fellowships,omitempty"`
	YouthFellowships                []int    `json:"youth_fellowships,omitempty"`
	ChildFellowships                []int    `json:"child_fellowships,omitempty"`
	Outreach                        []int    `json:"outreach,omitempty"`
	TrainingOrSeminars              []int    `json:"training_or_seminars,omitempty"`
	LeadershipConferences           []int    `json:"leadership_conferences,omitempty"`
	LeadershipTraining              []int    `json:"leadership_training,omitempty"`
	Others                          []int    `json:"others,omitempty"`
	FamilyDays                      []int    `json:"family_days,omitempty"`
	TithesAndOfferings              []int    `json:"tithes_and_offerings,omitempty"`
	HomeVisited                     []int    `json:"home_visited,omitempty"`
	BibleStudyOrGroupLed            []int    `json:"bible_study_or_group_led,omitempty"`
	SermonOrMessagePreached         []int    `json:"sermon_or_message_preached,omitempty"`
	PersonNewlyContacted            []int    `json:"person_newly_contacted,omitempty"`
	PersonFollowedUp                []int    `json:"person_followed_up,omitempty"`
	PersonLedToChrist               []int    `json:"person_led_to_christ,omitempty"`
	Names                           []string `json:"names,omitempty"`
	NarrativeReport                 string   `json:"narrative_report" validate:"required"`
	ChallengesAndProblemEncountered string   `json:"challenges_and_problem_encountered" validate:"required"`
	PrayerRequest                   string   `json:"prayer_request" validate:"required"`
	AverageAttendance               float64  `json:"average_attendance"`
}

func (request *ReportCreateRequest) Validate() error {
	if len(request.MonthOf) == 0 {
		return errors.New("month must not be empty")
	}

	if len(request.WorkerName) == 0 {
		return errors.New("worker mustt not be empty")
	}

	if len(request.AreaOfAssignment) == 0 {
		return errors.New("area of assignment must not be empty")
	}

	if len(request.NameOfChurch) == 0 {
		return errors.New("name of church must not be empty")
	}

	return nil
}
