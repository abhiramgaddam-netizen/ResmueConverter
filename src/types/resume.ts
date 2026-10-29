export interface ContactInfo {
  email?: string;
  phone?: string;
  location?: string;
  linkedin?: string;
  github?: string;
  portfolio?: string;
}

export interface WorkExperience {
  company: string;
  role: string;
  location?: string;
  startDate?: string;
  endDate?: string;
  bullets: string[];
}

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

export interface SkillCategory {
  category: string;
  skills: string[];
}

export interface AdditionalSection {
  title: string;
  items: string[];
}

export interface ResumeData {
  fullName: string;
  targetRole?: string;
  contactInfo: ContactInfo;
  professionalSummary?: string;
  skills: SkillCategory[];
  experience: WorkExperience[];
  education: EducationItem[];
  certifications?: string[];
  projects?: ProjectItem[];
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
