#from .cyschoolhousesuite import *
from . import section_creation, student, student_section
from .config import USER_SITE, get_sch_ref_df
from .simple_cysh import (get_object_df, get_object_fields, get_section_df,
                          get_staff_df, get_student_section_staff_df,
                          init_sf_session, object_reference, sf)

if USER_SITE == 'Chicago':
    from . import ia_assignment_chi
    from . import section_creation_chi
    from . import servicetrackers
    from . import tot_audit
    from . import tracker_mgmt
