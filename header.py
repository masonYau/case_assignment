class AnalystHeader:
    """分析师表的列名配置"""
    analyst_id = "Staff ID"  # 员工ID
    target_wip = "Optimal WIP"  # 目标产能
    current_wip = "Current WIP"  # 当前在手案件
    history_customers = "Served Clients"  # 历史客户 (假设是逗号分隔的字符串或列表)
    daily_productivity = "Last 3 month Productivity"  # 每周处理速度


class TaskHeader:
    """任务表的列名配置"""
    task_id = "Customer Number"  # 案件编号
    customer_id = "Customer Number"  # 客户名称
    due_date = "Anniversary Date"  # 截止日期
    effort = "Workload Units"  # 工作量权重
    market = "Market"
    legal_entity = "Legal Entity"
    ppm_cdd_source = "PPM CDD Source"
    CNAndLE = "CN & LE"
    risk_rating = "Risk Pyramid Risk Rating"


class ClosedReviewHeader:
    ReviewID = "Review ID"
    CustomerNumber = "Customer Number"
    LegalEntity = "Legal Entity"
    Market = "Market"
    Segment = "Segment"
    CustomerType = "Customer Type"
    ReviewReason = "Review Reason"
    RiskRating = "Risk Rating"
    RmNum = "RM Num"
    RmName = "RM Name"
    InitiatedDate = "Initiated Date"
    ApprovalCancelDate = "Approval/Cancel Date"
    ReviewStatus = "Review Status"
    LastCddReviewCompletionDate = "Last CDD Review Completion Date"
    Okyc = "OKYC"
    OkycInitiatedOtp = "OKYC Initiated (OTP)"
    LatestDcFinalisedById = "Latest DC Finalised by ID"
    LatestDcFinalisedByDate = "Latest DC Finalised Date"
    OkycLastUpdatedDate = "OKYC - Last Updated Date"
    OkycReviewStatus = "OKYC - Review Status"
    FirstClaimedID = "First Claimed ID"
    CancellationReason = "Cancellation Reason"

    StaffID = "Staff ID"
    StaffIDForHistory = "Staff ID For History"
    Team = "Team"
    GSC = "GSC"
    CaseCompletionMonth = "Case Completion Month"
    TurnAroundTime = "TurnAroundTime"
    RiskAdjFactor = "RiskAdjFactor"
    IDAndLE = "ID & LE"
    CNAndLE = "CN & LE"


class OpenReviewHeader:
    ReviewID = "Review ID"
    CustomerNumber = "Customer Number"
    CustomerName = "Customer Name"
    Market = "Market"
    LegalEntity = "Legal Entity"
    Segment = "Segment"
    RmNum = "RM Num"
    RmName = "RM Name"
    ReviewReason = "Review Reason"
    InitialRiskRatingCdd = "Initial Risk Rating (CDD)"
    RiskRating = "Risk Rating"
    Stage = "Stage"
    LatestAction = "Latest Action"
    StageClaimedDate = "Stage Claimed Date"
    TaskStatus = "Task Status"
    AssignedToRole = "Assigned to Role"
    AssignedToUser = "Assigned to User"
    DateOfLatestAction = "Date of latest action"
    StaffRole = "Staff Role"
    UserID = "User ID"
    InitiatedDate = "Initiated Date"
    FirstClaimedDate = "First Claimed Date"
    FirstClaimedID = "First Claimed ID"
    ApprovalCancelDate = "Approval/Cancel Date"
    LatestDcFinalisedById = "Latest DC Finalised by ID"
    LatestDcFinalisedDate = "Latest DC Finalised Date"
    TriggerStartDate = "Trigger Start Date"
    ActualRenewalDate = "Overdue Date"
    OkycDueDateBIB = "OKYC Due Date (BIB)"
    DueDateOTP = "Due date (OTP)"
    AssetApproxTotalAmt = "Asset Approx Total Amt"
    RiskPyramidLastReviewCompletedDate = "Risk Pyramid Last Review Completed Date"
    RiskPyramidRiskRating = "Risk Pyramid Risk Rating"
    CustomerIncorporationNumber = "Customer Incorporation Number"

    StaffID = "Staff ID"
    StaffIDForHistory = "Staff ID For History"
    Age = "Age"
    IsOverDue = "IsOverDue"
    Team = "Team"
    GSC = "GSC"
    SpecialCompletionUpdateDate = "SpecialCompletionUpdateDate"
    SpecialCompletionType = "SpecialCompletionType"
    IDAndLE = "ID & LE"
    CNAndLE = "CN & LE"


class MasterGroupHeader:
    CustomerNumber = "CIN"
    MasterGroupCode = "Mastergroup Code"


class RamHeader:
    ReviewID = "Real customer Id"
    CIN = "CIN"
    CustomerNumber = "Customer Number"
    CustomerName = "Customer Name"
    LOB = "LOB"
    LineOfBusiness = "Line of Business"
    NextCDDReviewDate = "Next CDD Review Date"
    DueDate = "Anniversary Date"


class StageType:
    DC = "DC"
    QC = "QC"
    APP = "APP"


class ReviewStatusType:
    ApprovalCompleted = "Approval Completed"
    Cancelled = "Cancelled"
    WIP = "WIP"
