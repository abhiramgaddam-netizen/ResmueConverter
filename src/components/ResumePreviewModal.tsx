'use client';

import React, { useState, useEffect } from 'react';
import {
  X,
  Download,
  Copy,
  Check,
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

  const displayName = (data?.fullName || candidateName || 'Candidate Profile').trim();
  const roleName = data?.targetRole;

  // Format full resume as structured plain text for copying
  const generateFormattedText = (): string => {
    if (!data) return extractedText || 'No resume data available.';

    const lines: string[] = [];
    lines.push(displayName);
    lines.push('--------------------------------------------------');

    const contactParts: string[] = [];
    if (!removePersonalInfo) {
      if (data.contactInfo?.email) contactParts.push(data.contactInfo.email);
      if (data.contactInfo?.phone) contactParts.push(data.contactInfo.phone);
      if (data.contactInfo?.location) contactParts.push(data.contactInfo.location);
      if (data.contactInfo?.linkedin) contactParts.push(data.contactInfo.linkedin);
      if (data.contactInfo?.github) contactParts.push(data.contactInfo.github);
    }
    if (contactParts.length > 0) {
      lines.push(contactParts.join('   |   '));
      lines.push('');
    }

    // 1. Professional Summary
    const summaryItems = Array.isArray(data.professionalSummary)
      ? data.professionalSummary
      : typeof data.professionalSummary === 'string' && data.professionalSummary.trim()
      ? data.professionalSummary.split('\n').filter(Boolean)
      : [];

    if (summaryItems.length > 0) {
      lines.push('Professional Summary:');
      summaryItems.forEach((item) => lines.push(`•\t${item}`));
      lines.push('');
    }

    // Specialized Skills
    if (data.specializedSkills && data.specializedSkills.length > 0) {
      data.specializedSkills.forEach((g) => {
        if (g.title) lines.push(g.title);
        (g.bullets || []).forEach((b) => lines.push(`•\t${b}`));
        lines.push('');
      });
    }

    // 2. Education
    if (data.education && data.education.length > 0) {
      lines.push('Education:');
      data.education.forEach((edu) => {
        const line = [
          edu.degree,
          edu.fieldOfStudy ? `in ${edu.fieldOfStudy}` : '',
          edu.institution ? `from ${edu.institution}` : '',
          edu.endDate ? `(${edu.endDate})` : '',
        ]
          .filter(Boolean)
          .join(' ')
          .trim();
        if (line) lines.push(`•\t${line}`);
      });
      lines.push('');
    }

    // 3. Technical Skills
    if (data.skills && data.skills.length > 0) {
      lines.push('Technical Skills:');
      data.skills.forEach((cat) => {
        lines.push(`${cat.category}:\t${cat.skills.join(', ')}`);
      });
      lines.push('');
    }

    // 4. Awards and Achievements
    if (data.awards && data.awards.length > 0) {
      lines.push('Awards and Achievements:');
      data.awards.forEach((award) => lines.push(`•\t${award}`));
      lines.push('');
    }

    // 5. Professional Experience
    if (data.experience && data.experience.length > 0) {
      lines.push('Professional Experience:');
      data.experience.forEach((exp, idx) => {
        const projTitle = exp.projectNumber || `Project ${idx + 1}`;
        lines.push(projTitle);
        const duration = [exp.startDate, exp.endDate].filter(Boolean).join(' to ');
        if (exp.company) lines.push(`Client:\t\t${exp.company}`);
        if (exp.role) lines.push(`Role:\t\t${exp.role}`);
        if (duration) lines.push(`Duration:\t${duration}`);
        if (exp.environment) lines.push(`Environment:\t${exp.environment}`);

        if (exp.workingEnvironment && exp.workingEnvironment.length > 0) {
          lines.push('Working Environment:');
          exp.workingEnvironment.forEach((we) => lines.push(`•\t${we}`));
        }

        if (exp.bullets && exp.bullets.length > 0) {
          lines.push('Responsibilities:');
          exp.bullets.forEach((b) => lines.push(`•\t${b}`));
        }
        lines.push('');
      });
    }

    // 6. Certifications
    if (data.certifications && data.certifications.length > 0) {
      lines.push('Professional Certifications:');
      data.certifications.forEach((cert) => lines.push(`•\t${cert}`));
      lines.push('');
    }

    return lines.join('\n');
  };

  const handleCopy = () => {
    const textToCopy = generateFormattedText();
    navigator.clipboard.writeText(textToCopy);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  // Build Contact Items bar
  const contactItems: string[] = [];
  if (!removePersonalInfo) {
    if (data?.contactInfo?.email) contactItems.push(data.contactInfo.email);
    if (data?.contactInfo?.phone) contactItems.push(data.contactInfo.phone);
    if (data?.contactInfo?.location) contactItems.push(data.contactInfo.location);
    if (data?.contactInfo?.linkedin) contactItems.push(data.contactInfo.linkedin);
    if (data?.contactInfo?.github) contactItems.push(data.contactInfo.github);
  }

  // Summary items
  const summaryItems = Array.isArray(data?.professionalSummary)
    ? data.professionalSummary
    : typeof data?.professionalSummary === 'string' && data.professionalSummary.trim()
    ? data.professionalSummary.split('\n').filter(Boolean)
    : [];

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-2 sm:p-4 bg-slate-900/60 backdrop-blur-xs animate-fadeIn"
      onClick={onClose}
    >
      <div
        className="bg-white rounded-xl shadow-2xl border border-slate-200 flex flex-col w-full max-w-5xl h-[94vh] overflow-hidden"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Modal Top Control Bar */}
        <div className="flex items-center justify-between px-6 py-3.5 border-b border-slate-200 bg-white">
          <div className="flex items-center space-x-3 min-w-0">
            <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 flex-shrink-0 animate-pulse" />
            <div className="min-w-0">
              <h2 className="text-sm font-semibold text-slate-900 truncate">
                {displayName}
              </h2>
              <p className="text-[11px] text-slate-500 truncate">
                {originalFilename} &bull; Classic Tecnics Format
              </p>
            </div>
          </div>

          <div className="flex items-center space-x-2">
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
            className="w-full max-w-[850px] bg-white border border-slate-300 shadow-xl rounded-xs p-10 sm:p-14 my-auto text-black"
            style={{
              fontFamily: '"Times New Roman", Times, serif',
              minHeight: '1050px',
            }}
          >
            {/* 1. EXACT WORD HEADER: Candidate Name on Left + Official TECNICS Logo on Right */}
            <div className="flex items-center justify-between pb-2 border-b border-black">
              <div className="min-w-0 pr-4">
                <h1 className="text-xl sm:text-2xl font-bold text-black tracking-normal truncate">
                  {headerType === 'role' && roleName ? roleName : displayName}
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

            {/* Sub-Header: Role + Contact info (if present) */}
            {(roleName || contactItems.length > 0) && (
              <div className="mt-3 mb-4">
                {roleName && headerType !== 'role' && (
                  <div className="font-bold text-base text-black mb-1">
                    {roleName}
                  </div>
                )}
                {contactItems.length > 0 && (
                  <div className="text-xs text-slate-600">
                    {contactItems.join('   |   ')}
                  </div>
                )}
              </div>
            )}

            {/* Fallback if no structured data available */}
            {!data && (
              <div className="mt-6">
                <div className="font-bold text-sm text-black mb-2">
                  EXTRACTED SOURCE TEXT:
                </div>
                <pre className="text-xs font-mono text-slate-800 whitespace-pre-wrap leading-relaxed">
                  {extractedText || 'No resume data available.'}
                </pre>
              </div>
            )}

            {/* SECTION 1: PROFESSIONAL SUMMARY */}
            {summaryItems.length > 0 && (
              <div className="mb-4">
                <h2 className="text-[14px] font-bold text-black mb-1.5">
                  Professional Summary:
                </h2>
                <ul className="space-y-1 pl-5 text-[12.5px] text-black">
                  {summaryItems.map((item, idx) => (
                    <li key={idx} className="list-disc leading-relaxed pl-1">
                      {item}
                    </li>
                  ))}
                </ul>
              </div>
            )}

            {/* SPECIALIZED SKILLS (Migration Projects, SKILLS@SAP R/3, etc.) */}
            {data?.specializedSkills && data.specializedSkills.length > 0 && (
              <div className="mb-4 space-y-3">
                {data.specializedSkills.map((group, gIdx) => (
                  <div key={gIdx}>
                    <h3 className="text-[12.5px] font-bold text-black mb-1">
                      {group.title}
                    </h3>
                    {group.bullets && group.bullets.length > 0 && (
                      <ul className="space-y-1 pl-5 text-[12.5px] text-black">
                        {group.bullets.map((b, bIdx) => (
                          <li key={bIdx} className="list-disc leading-relaxed pl-1">
                            {b}
                          </li>
                        ))}
                      </ul>
                    )}
                  </div>
                ))}
              </div>
            )}

            {/* SECTION 2: EDUCATION */}
            {data?.education && data.education.length > 0 && (
              <div className="mb-4">
                <h2 className="text-[14px] font-bold text-black mb-1.5">
                  Education:
                </h2>
                <ul className="space-y-1 pl-5 text-[12.5px] text-black">
                  {data.education.map((edu, idx) => {
                    const line = [
                      edu.degree,
                      edu.fieldOfStudy ? `in ${edu.fieldOfStudy}` : '',
                      edu.institution ? `from ${edu.institution}` : '',
                      edu.endDate ? `(${edu.endDate})` : '',
                    ]
                      .filter(Boolean)
                      .join(' ')
                      .trim();
                    return (
                      <li key={idx} className="list-disc leading-relaxed pl-1">
                        {line}
                      </li>
                    );
                  })}
                </ul>
              </div>
            )}

            {/* SECTION 3: TECHNICAL SKILLS */}
            {data?.skills && data.skills.length > 0 && (
              <div className="mb-4">
                <h2 className="text-[14px] font-bold text-black mb-1.5">
                  Technical Skills:
                </h2>
                <div className="space-y-1 text-[12.5px] text-black">
                  {data.skills.map((cat, idx) => (
                    <div key={idx} className="grid grid-cols-12 gap-2">
                      <span className="col-span-4 sm:col-span-3 text-black font-normal">
                        {cat.category.endsWith(':') ? cat.category : `${cat.category}:`}
                      </span>
                      <span className="col-span-8 sm:col-span-9 text-black">
                        {cat.skills.join(', ')}
                      </span>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* SECTION 4: AWARDS AND ACHIEVEMENTS */}
            {data?.awards && data.awards.length > 0 && (
              <div className="mb-4">
                <h2 className="text-[14px] font-bold text-black mb-1.5">
                  Awards and Achievements:
                </h2>
                <ul className="space-y-1 pl-5 text-[12.5px] text-black">
                  {data.awards.map((award, idx) => (
                    <li key={idx} className="list-disc leading-relaxed pl-1">
                      {award}
                    </li>
                  ))}
                </ul>
              </div>
            )}

            {/* SECTION 5: PROFESSIONAL EXPERIENCE */}
            {data?.experience && data.experience.length > 0 && (
              <div className="mb-4">
                <h2 className="text-[14px] font-bold text-black mb-2">
                  Professional Experience:
                </h2>
                <div className="space-y-4">
                  {data.experience.map((exp, idx) => {
                    const projTitle = exp.projectNumber || `Project ${idx + 1}`;
                    const duration = [exp.startDate, exp.endDate].filter(Boolean).join(' to ');

                    return (
                      <div key={idx} className="text-[12.5px] text-black space-y-1.5">
                        <div className="font-bold text-[13px] text-black">
                          {projTitle}
                        </div>

                        {/* Key-Value Block */}
                        <div className="space-y-0.5">
                          {exp.company && (
                            <div className="grid grid-cols-12 gap-2">
                              <span className="col-span-3 sm:col-span-2 font-bold text-black">
                                Client:
                              </span>
                              <span className="col-span-9 sm:col-span-10 text-black">
                                {exp.company}
                              </span>
                            </div>
                          )}
                          {exp.role && (
                            <div className="grid grid-cols-12 gap-2">
                              <span className="col-span-3 sm:col-span-2 font-bold text-black">
                                Role:
                              </span>
                              <span className="col-span-9 sm:col-span-10 text-black">
                                {exp.role}
                              </span>
                            </div>
                          )}
                          {duration && (
                            <div className="grid grid-cols-12 gap-2">
                              <span className="col-span-3 sm:col-span-2 font-bold text-black">
                                Duration:
                              </span>
                              <span className="col-span-9 sm:col-span-10 text-black">
                                {duration}
                              </span>
                            </div>
                          )}
                          {exp.environment && (
                            <div className="grid grid-cols-12 gap-2">
                              <span className="col-span-3 sm:col-span-2 font-bold text-black">
                                Environment:
                              </span>
                              <span className="col-span-9 sm:col-span-10 text-black">
                                {exp.environment}
                              </span>
                            </div>
                          )}
                        </div>

                        {/* Working Environment */}
                        {exp.workingEnvironment && exp.workingEnvironment.length > 0 && (
                          <div className="mt-2">
                            <div className="font-bold text-black mb-1">
                              Working Environment:
                            </div>
                            <ul className="space-y-1 pl-5 text-black">
                              {exp.workingEnvironment.map((we, weIdx) => (
                                <li key={weIdx} className="list-disc leading-relaxed pl-1">
                                  {we}
                                </li>
                              ))}
                            </ul>
                          </div>
                        )}

                        {/* Responsibilities */}
                        {exp.bullets && exp.bullets.length > 0 && (
                          <div className="mt-2">
                            <div className="font-bold text-black mb-1">
                              Responsibilities:
                            </div>
                            <ul className="space-y-1 pl-5 text-black">
                              {exp.bullets.map((b, bIdx) => (
                                <li key={bIdx} className="list-disc leading-relaxed pl-1">
                                  {b}
                                </li>
                              ))}
                            </ul>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {/* SECTION 6: PROFESSIONAL CERTIFICATIONS */}
            {data?.certifications && data.certifications.length > 0 && (
              <div className="mb-4">
                <h2 className="text-[14px] font-bold text-black mb-1.5">
                  Professional Certifications:
                </h2>
                <ul className="space-y-1 pl-5 text-[12.5px] text-black">
                  {data.certifications.map((cert, idx) => (
                    <li key={idx} className="list-disc leading-relaxed pl-1">
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
