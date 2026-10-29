export interface ContactInfo {
  email?: string;
  phone?: string;
  location?: string;
  linkedin?: string;
  github?: string;
  portfolio?: string;
}

export interface SpecializedSkillGroup {
  title: string;      // e.g., "Migration Projects", "SKILLS@SAP R/3", "Net Weaver Skills", "Hana Skills"
  bullets: string[];
}

export interface TechnicalSkillCategory {
  category: string;   // e.g., "Operating Systems", "SAP Software", "Database"
  skills: string[];
}

// Alias for backwards compatibility
export type SkillCategory = TechnicalSkillCategory;

export interface ProjectExperience {
  projectNumber?: string;        // e.g., "Project 1", "Project 2"
  company: string;               // Client Name
  role: string;                  // Role Title
  startDate: string;             // Duration (or start date)
  endDate?: string;
  location?: string;
  environment?: string;          // Full environment string
  workingEnvironment?: string[]; // Bullet points under "Working Environment:" if present
  bullets: string[];             // Bullet points under "Responsibilities:"
}

// Alias for backwards compatibility
export type WorkExperience = ProjectExperience;

export interface EducationItem {
  institution: string;
  degree: string;
  fieldOfStudy?: string;
  startDate?: string;
  endDate?: string;
  details?: string[];
}

export interface ProjectItem {
  name: string;
  technologies?: string;
  description?: string;
  bullets?: string[];
}

export interface AdditionalSection {
  title: string;
  items: string[];
}

export interface ResumeData {
  fullName: string;
  targetRole: string;
  contactInfo?: {
    email?: string;
    phone?: string;
    location?: string;
    linkedin?: string;
    github?: string;
  };
  professionalSummary: string | string[]; // Supports bullet array or string
  specializedSkills?: SpecializedSkillGroup[]; // Captures "Migration Projects", "SKILLS@SAP R/3", etc.
  education: EducationItem[];
  skills: TechnicalSkillCategory[];
  awards?: string[];
  experience: ProjectExperience[];
  projects?: any[];
  certifications?: string[];
  additionalSections?: AdditionalSection[];
}

export interface ConversionResult {
  id: string;
  fileName: string;
  candidateName: string;
  status: 'pending' | 'processing' | 'completed' | 'error';
  docxBase64?: string;
  error?: string;
  data?: ResumeData;
}
