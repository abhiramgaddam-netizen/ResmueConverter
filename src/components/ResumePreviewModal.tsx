'use client';

import React, { useState, useEffect } from 'react';
import {
  X,
  Download,
  Copy,
  Check,
  FileText,
  Printer,
} from 'lucide-react';
import { ResumeData } from '@/types/resume';
import { TECNICS_LOGO_DATA_URI } from '@/lib/logo';

interface ResumePreviewModalProps {
  isOpen: boolean;
  onClose: () => void;
  candidateName: string;
  originalFilename: string;
  data?: ResumeData;
  extractedText?: string;
  docxBase64?: string;
  onDownloadDocx?: () => void;
  removePersonalInfo?: boolean;
  headerType?: 'role' | 'name';
}

export default function ResumePreviewModal({
  isOpen,
  onClose,
  candidateName,
  originalFilename,
  data,
  extractedText,
  docxBase64,
  onDownloadDocx,
  removePersonalInfo = true,
  headerType = 'name',
}: ResumePreviewModalProps) {
  const [copied, setCopied] = useState<boolean>(false);

  // Close on Escape key press
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    if (isOpen) {
      document.body.style.overflow = 'hidden';
      window.addEventListener('keydown', handleKeyDown);
    }
    return () => {
      document.body.style.overflow = 'unset';
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [isOpen, onClose]);

  if (!isOpen) return null;

  // Format full resume as structured plain text for copying
  const generateFormattedText = (): string => {
    if (!data) return extractedText || 'No resume data available.';

    const lines: string[] = [];
    lines.push((data.fullName || candidateName).toUpperCase());
    lines.push('--------------------------------------------------');
    if (data.targetRole) lines.push(data.targetRole);

    const contactParts: string[] = [];
    if (!removePersonalInfo) {
      if (data.contactInfo?.email) contactParts.push(data.contactInfo.email);
      if (data.contactInfo?.phone) contactParts.push(data.contactInfo.phone);
      if (data.contactInfo?.location) contactParts.push(data.contactInfo.location);
      if (data.contactInfo?.linkedin) contactParts.push(data.contactInfo.linkedin);
      if (data.contactInfo?.github) contactParts.push(data.contactInfo.github);
    }
    if (contactParts.length > 0) lines.push(contactParts.join('   |   '));
    lines.push('');

    if (data.professionalSummary) {
      lines.push('PROFESSIONAL SUMMARY');
      lines.push(data.professionalSummary);
      lines.push('');
    }

    if (data.skills && data.skills.length > 0) {
      lines.push('TECHNICAL SKILLS & COMPETENCIES');
      data.skills.forEach((cat) => {
        lines.push(`${cat.category}: ${cat.skills.join(', ')}`);
      });
      lines.push('');
    }

    if (data.experience && data.experience.length > 0) {
      lines.push('PROFESSIONAL EXPERIENCE');
      data.experience.forEach((exp) => {
        lines.push(exp.role);
        const meta = [exp.company, exp.location, [exp.startDate, exp.endDate].filter(Boolean).join(' - ')].filter(Boolean).join('  |  ');
        if (meta) lines.push(meta);
        exp.bullets.forEach((b) => lines.push(`•  ${b}`));
        lines.push('');
      });
    }

    if (data.projects && data.projects.length > 0) {
      lines.push('KEY PROJECTS & ENGAGEMENTS');
      data.projects.forEach((proj) => {
        lines.push(`${proj.name} ${proj.technologies ? `(${proj.technologies})` : ''}`);
        if (proj.description) lines.push(proj.description);
        if (proj.bullets) proj.bullets.forEach((b) => lines.push(`•  ${b}`));
        lines.push('');
      });
    }

    if (data.education && data.education.length > 0) {
      lines.push('EDUCATION & CREDENTIALS');
      data.education.forEach((edu) => {
        const dates = [edu.startDate, edu.endDate].filter(Boolean).join(' - ');
        lines.push(`${edu.institution}${dates ? `  (${dates})` : ''}`);
        const degree = [edu.degree, edu.fieldOfStudy].filter(Boolean).join(' in ');
        if (degree) lines.push(degree);
        lines.push('');
      });
    }

    if (data.certifications && data.certifications.length > 0) {
      lines.push('PROFESSIONAL CERTIFICATIONS');
      data.certifications.forEach((cert) => {
        lines.push(`•  ${cert}`);
      });
    }

    return lines.join('\n');
  };

  const handleCopy = () => {
    const textToCopy = generateFormattedText();
    navigator.clipboard.writeText(textToCopy);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const displayName = (data?.fullName || candidateName || 'Candidate Profile').trim();
  const roleName = data?.targetRole;

  // Build Contact Items bar exactly like Word doc
  const contactItems: string[] = [];
  if (!removePersonalInfo) {
    if (data?.contactInfo?.email) contactItems.push(data.contactInfo.email);
    if (data?.contactInfo?.phone) contactItems.push(data.contactInfo.phone);
    if (data?.contactInfo?.location) contactItems.push(data.contactInfo.location);
    if (data?.contactInfo?.linkedin) contactItems.push(data.contactInfo.linkedin);
    if (data?.contactInfo?.github) contactItems.push(data.contactInfo.github);
  }

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-2 sm:p-4 bg-slate-900/60 backdrop-blur-xs animate-fadeIn"
      onClick={onClose}
    >
      <div
        className="bg-slate-100 rounded-xl shadow-2xl border border-slate-300 w-full max-w-5xl h-[94vh] flex flex-col overflow-hidden text-slate-900"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Top Control Bar */}
        <div className="px-5 py-3 border-b border-slate-200 bg-white flex items-center justify-between gap-3 flex-shrink-0">
          <div className="flex items-center space-x-2.5 min-w-0">
            <div className="w-8 h-8 rounded bg-blue-50 border border-blue-200 flex items-center justify-center text-blue-700 flex-shrink-0">
              <FileText className="w-4 h-4" />
            </div>
            <div className="min-w-0">
              <div className="flex items-center space-x-2">
                <h3 className="text-sm font-bold text-slate-900 truncate">
                  {displayName}
                </h3>
                <span className="text-[10px] font-semibold px-2 py-0.5 rounded bg-blue-50 text-blue-700 border border-blue-200 uppercase tracking-wide hidden sm:inline-block">
                  Word Document Layout
                </span>
                {removePersonalInfo && (
                  <span className="text-[10px] font-semibold px-2 py-0.5 rounded bg-emerald-50 text-emerald-700 border border-emerald-200 uppercase tracking-wide hidden sm:inline-block">
                    Personal Details Removed
                  </span>
                )}
              </div>
              <p className="text-[11px] text-slate-500 truncate">Source: {originalFilename}</p>
            </div>
          </div>

          {/* Action Buttons */}
          <div className="flex items-center space-x-2 flex-shrink-0">
            {/* Copy Button */}
            <button
              onClick={handleCopy}
              className="px-3 py-1.5 rounded-lg text-xs font-medium border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 flex items-center space-x-1.5 transition-colors shadow-xs"
              title="Copy formatted text"
            >
              {copied ? (
                <>
                  <Check className="w-3.5 h-3.5 text-emerald-600" />
                  <span className="text-emerald-700 font-semibold">Copied!</span>
                </>
              ) : (
                <>
                  <Copy className="w-3.5 h-3.5 text-slate-500" />
                  <span>Copy</span>
                </>
              )}
            </button>

            {/* Print Button */}
            <button
              onClick={() => window.print()}
              className="px-3 py-1.5 rounded-lg text-xs font-medium border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 hidden sm:flex items-center space-x-1.5 transition-colors shadow-xs"
              title="Print document preview"
            >
              <Printer className="w-3.5 h-3.5 text-slate-500" />
              <span>Print</span>
            </button>

            {/* Download DOCX Button */}
            {docxBase64 && onDownloadDocx && (
              <button
                onClick={onDownloadDocx}
                className="bg-blue-600 hover:bg-blue-700 text-white px-3.5 py-1.5 rounded-lg text-xs font-semibold flex items-center space-x-1.5 transition-colors shadow-xs"
                title="Download standardized Word Document (.docx)"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Download .docx</span>
              </button>
            )}

            {/* Close Button */}
            <button
              onClick={onClose}
              className="p-1.5 rounded-lg text-slate-400 hover:text-slate-700 hover:bg-slate-100 transition-colors ml-1"
              title="Close (Esc)"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* Modal Scrollable Canvas: Renders Exact Word Document Page Layout */}
        <div className="flex-1 overflow-y-auto p-4 sm:p-8 bg-slate-200/70 flex justify-center">
          <div
            className="w-full max-w-[850px] bg-white border border-slate-300 shadow-xl rounded-xs p-10 sm:p-14 my-auto text-slate-900"
            style={{
              fontFamily: 'Calibri, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
              minHeight: '1050px',
            }}
          >
            {/* 1. EXACT WORD HEADER: Candidate Name or Role on Left + Official TECNICS Logo on Right */}
            <div className="flex items-center justify-between pb-3 border-b border-slate-300">
              <div className="min-w-0 pr-4">
                <h1 className="text-xl sm:text-2xl font-bold tracking-widest text-[#475569] uppercase truncate">
                  {headerType === 'role' && roleName ? roleName.toUpperCase() : displayName.toUpperCase()}
                </h1>
              </div>

              {/* Official Tecnics Logo Image */}
              <div className="flex-shrink-0">
                <img
                  src={TECNICS_LOGO_DATA_URI}
                  alt="Tecnics Logo"
                  className="h-8 sm:h-9 object-contain"
                />
              </div>
            </div>

            {/* 2. SUB-HEADER: Role (in Blue) + Single Line Contact Info */}
            <div className="mt-4 mb-5">
              {roleName && (
                <div className="text-[#0284c7] font-bold text-[15px] sm:text-base leading-tight mb-1">
                  {roleName}
                </div>
              )}

              {contactItems.length > 0 && (
                <div className="text-[11.5px] sm:text-xs text-[#64748b] leading-relaxed">
                  {contactItems.join('   |   ')}
                </div>
              )}
            </div>

            {/* Fallback if no structured data available yet */}
            {!data && (
              <div className="mt-6">
                <div className="text-[#0284c7] font-bold text-xs uppercase tracking-wider border-b border-[#0284c7] pb-0.5 mb-2">
                  EXTRACTED SOURCE TEXT
                </div>
                <pre className="text-xs font-mono text-slate-700 whitespace-pre-wrap leading-relaxed">
                  {extractedText || 'No resume data available.'}
                </pre>
              </div>
            )}

            {/* 3. SECTION 1: PROFESSIONAL SUMMARY */}
            {data?.professionalSummary && (
              <div className="mb-5">
                <div className="text-[#0284c7] font-bold text-xs sm:text-[13px] uppercase tracking-wider border-b border-[#0284c7] pb-0.5 mb-2">
                  PROFESSIONAL SUMMARY
                </div>
                <p className="text-xs sm:text-[13px] text-[#334155] leading-relaxed text-justify">
                  {data.professionalSummary}
                </p>
              </div>
            )}

            {/* 4. SECTION 2: TECHNICAL SKILLS & COMPETENCIES */}
            {data?.skills && data.skills.length > 0 && (
              <div className="mb-5">
                <div className="text-[#0284c7] font-bold text-xs sm:text-[13px] uppercase tracking-wider border-b border-[#0284c7] pb-0.5 mb-2">
                  TECHNICAL SKILLS &amp; COMPETENCIES
                </div>
                <div className="space-y-1">
                  {data.skills.map((cat, idx) => (
                    <div key={idx} className="text-xs sm:text-[13px] text-[#334155] leading-relaxed">
                      <strong className="text-[#0f172a] font-bold">{cat.category}: </strong>
                      <span>{cat.skills.join(', ')}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* 5. SECTION 3: PROFESSIONAL EXPERIENCE */}
            {data?.experience && data.experience.length > 0 && (
              <div className="mb-5">
                <div className="text-[#0284c7] font-bold text-xs sm:text-[13px] uppercase tracking-wider border-b border-[#0284c7] pb-0.5 mb-2.5">
                  PROFESSIONAL EXPERIENCE
                </div>
                <div className="space-y-4">
                  {data.experience.map((exp, idx) => {
                    const meta = [
                      exp.company,
                      exp.location,
                      [exp.startDate, exp.endDate].filter(Boolean).join(' - '),
                    ]
                      .filter(Boolean)
                      .join('  |  ');

                    return (
                      <div key={idx} className="leading-snug">
                        <div className="text-xs sm:text-[13px] font-bold text-[#0f172a]">
                          {exp.role}
                        </div>
                        {meta && (
                          <div className="text-xs text-[#0284c7] italic mb-1">
                            {meta}
                          </div>
                        )}
                        {exp.bullets && exp.bullets.length > 0 && (
                          <ul className="space-y-1 pl-4 text-xs sm:text-[13px] text-[#334155]">
                            {exp.bullets.map((b, bIdx) => (
                              <li key={bIdx} className="list-disc leading-relaxed pl-1 marker:text-[#0284c7]">
                                {b}
                              </li>
                            ))}
                          </ul>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {/* 6. SECTION 4: KEY PROJECTS & ENGAGEMENTS */}
            {data?.projects && data.projects.length > 0 && (
              <div className="mb-5">
                <div className="text-[#0284c7] font-bold text-xs sm:text-[13px] uppercase tracking-wider border-b border-[#0284c7] pb-0.5 mb-2.5">
                  KEY PROJECTS &amp; ENGAGEMENTS
                </div>
                <div className="space-y-3.5">
                  {data.projects.map((proj, idx) => (
                    <div key={idx}>
                      <div className="text-xs sm:text-[13px] text-[#0f172a]">
                        <strong className="font-bold">{proj.name}</strong>
                        {proj.technologies && (
                          <span className="italic text-[#64748b]"> ({proj.technologies})</span>
                        )}
                      </div>
                      {proj.description && (
                        <p className="text-xs sm:text-[13px] text-[#334155] leading-relaxed mb-1 mt-0.5">
                          {proj.description}
                        </p>
                      )}
                      {proj.bullets && proj.bullets.length > 0 && (
                        <ul className="space-y-1 pl-4 text-xs sm:text-[13px] text-[#334155]">
                          {proj.bullets.map((b, bIdx) => (
                            <li key={bIdx} className="list-disc leading-relaxed pl-1 marker:text-[#0284c7]">
                              {b}
                            </li>
                          ))}
                        </ul>
                      )}
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* 7. SECTION 5: EDUCATION & CREDENTIALS */}
            {data?.education && data.education.length > 0 && (
              <div className="mb-5">
                <div className="text-[#0284c7] font-bold text-xs sm:text-[13px] uppercase tracking-wider border-b border-[#0284c7] pb-0.5 mb-2">
                  EDUCATION &amp; CREDENTIALS
                </div>
                <div className="space-y-2">
                  {data.education.map((edu, idx) => {
                    const dates = [edu.startDate, edu.endDate].filter(Boolean).join(' - ');
                    const degree = [edu.degree, edu.fieldOfStudy].filter(Boolean).join(' in ');

                    return (
                      <div key={idx} className="text-xs sm:text-[13px]">
                        <div className="font-bold text-[#0f172a]">
                          {edu.institution}
                          {dates && <span className="font-normal text-[#64748b]">  ({dates})</span>}
                        </div>
                        {degree && (
                          <div className="text-[#334155]">{degree}</div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {/* 8. SECTION 6: PROFESSIONAL CERTIFICATIONS */}
            {data?.certifications && data.certifications.length > 0 && (
              <div className="mb-5">
                <div className="text-[#0284c7] font-bold text-xs sm:text-[13px] uppercase tracking-wider border-b border-[#0284c7] pb-0.5 mb-2">
                  PROFESSIONAL CERTIFICATIONS
                </div>
                <ul className="space-y-1 pl-4 text-xs sm:text-[13px] text-[#334155]">
                  {data.certifications.map((cert, idx) => (
                    <li key={idx} className="list-disc leading-relaxed pl-1 marker:text-[#0284c7]">
                      {cert}
                    </li>
                  ))}
                </ul>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
