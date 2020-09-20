#from .cyschoolhousesuite import *
from . import section_creation, student, student_section
from .config import USER_SITE
from .simple_cysh import (get_object_df, get_object_fields, get_section_df,
                          get_staff_df, get_student_df,
                          get_student_section_staff_df, init_sf_session,
                          object_reference, sf)
from .utils import map_sharepoint_drive, get_sch_ref_df

if USER_SITE.lower() == 'chicago':
    from . import chi_ia_assignment
    from . import chi_section_creation
    from . import chi_thrive_datashare
    from .tot_audit import ToTAudit
    from .trackers import (AttendanceTracker, CoachingLog, 
                           LeadershipTracker, WeeklyServiceTracker)

map_sharepoint_drive()
