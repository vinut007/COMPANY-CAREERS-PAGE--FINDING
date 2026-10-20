import sys
sys.path.insert(0, r"C:\Users\vinutshreedhar.n\Downloads\company careers")
from find_careers import JOBS_LIST_TEXT_RE

tests = [
    "View Jobs", "Search Jobs", "Find Jobs", "Browse Jobs", "Browse Open Positions",
    "View Open Positions", "Current Openings", "Current Opportunities", "All Jobs",
    "See All Jobs", "Explore Open Roles", "Search Openings", "Job Listings",
    "Job Opportunities", "Available Positions", "Open Positions", "Find Your Role",
    "Explore Jobs", "See Jobs", "See Openings", "View All Jobs", "View All Opportunities",
    "View All Openings", "Explore All Jobs", "Explore Job Opportunities",
    "Search Career Opportunities", "Job Search", "Career Search", "Open Roles",
    "Available Jobs", "Latest Jobs", "Latest Openings", "New Opportunities",
    "Recently Posted Jobs", "Featured Jobs", "Engineering Jobs",
    "Software Engineering Careers", "Technology Careers", "Tech Jobs",
    "Backend Developer Jobs", "Developer Opportunities", "IT Careers",
    "Internship Opportunities", "Internships", "Graduate Opportunities", "Graduate Jobs",
    "Early Careers", "University Recruiting", "Campus Recruiting", "Student Opportunities",
    "Entry-Level Jobs", "Experienced Professionals", "Professional Careers",
    "Remote Jobs", "Work From Home Jobs", "Contract Opportunities",
    "Full-Time Jobs", "Part-Time Jobs", "Jobs by Location", "Jobs by Department",
    "Jobs by Team", "Teams Hiring", "Departments Hiring",
    "Join Our Engineering Team", "Join Our Tech Team", "Build Your Career",
    "Start Your Career", "Grow With Us", "Make an Impact", "Find Your Next Role",
    "Your Career Starts Here", "Apply Now", "Apply Here", "Apply Today",
    "View Job Details", "View Details", "Job Details", "Read More", "Learn More",
    "Submit Your Resume", "Upload Resume", "Register Your Interest",
    "Sign Up for Job Alerts", "Create Job Alert", "Get Job Alerts",
    "Talent Community", "Join Our Talent Network", "Join Our Talent Community",
    "Connect With Recruiters", "Meet Our Team", "Hiring Now", "We're Hiring",
    "Now Hiring", "Current Vacancies", "Vacancies", "View Vacancies", "Search Vacancies",
    "Browse Vacancies", "Explore Vacancies", "Vacancy Listings", "Find a Position",
    "Browse Positions", "View Positions", "Search Positions", "Employment Opportunities",
    "Work Opportunities", "Recruitment", "Recruitment Opportunities",
    "Career Portal", "Job Portal", "Applicant Portal", "Candidate Portal",
    "External Careers", "Visit Careers Site", "Visit Job Site", "Go to Careers",
    "Go to Jobs", "See Open Roles", "View Available Roles", "Explore Available Positions",
    "Browse Career Opportunities", "Find Open Positions", "Search Open Roles",
    "View All Positions", "View All Vacancies",
    "Careers", "Career", "Career Opportunities", "Explore Careers", "Explore Opportunities",
    "View Opportunities", "Discover Opportunities", "Join Our Team", "Work With Us",
    "Life at Company", "Careers at Company", "Discover Careers", "Find Your Future",
    "See OPEN POSITIONS", "Job Openings", "Our Openings",
]

miss = [t for t in tests if not JOBS_LIST_TEXT_RE.search(t)]
print(f"Matched: {len(tests) - len(miss)}/{len(tests)}")
print("MISSED:", miss if miss else "none")
