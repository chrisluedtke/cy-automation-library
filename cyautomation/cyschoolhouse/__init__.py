#from .cyschoolhousesuite import *
from .config import get_sch_ref_df, USER_SITE
from .simple_cysh import *

from . import section_creation
from . import student
from . import student_section

if USER_SITE == 'Chicago':
    from . import ia_assignment_chi
    from . import section_creation_chi
    from . import servicetrackers
    from . import tot_audit
    from . import tracker_mgmt
