function sendReportMails_(report) {
  const messages = [];

  try {
    const adminTo = normalizeMailRecipients_(getSettingValue_(SETTING_MAIL_TO));
    if (adminTo) {
      sendAdminDailyReportMail_(adminTo, report);
      messages.push('管理者へメール送信しました。');
    } else {
      messages.push('管理者メール送信先が未設定のため、管理者メールは送信していません。');
    }
  } catch (e) {
    console.error(e);
    messages.push('管理者メール送信でエラーが発生しました。');
  }

  try {
    const userEmail = normalizeMailRecipients_(report.userEmail || '');
    if (userEmail) {
      sendUserDailyReportMail_(userEmail, report);
      messages.push('入力者へ確認メールを送信しました。');
    } else {
      messages.push('入力者メールアドレス未設定のため、確認メールは送信していません。');
    }
  } catch (e) {
    console.error(e);
    messages.push('入力者向け確認メール送信でエラーが発生しました。');
  }

  return {
    message: messages.join(' '),
  };
}

function sendAdminDailyReportMail_(mailTo, report) {
  sendMail_(mailTo, buildAdminMailSubject_(report), buildDailyReportMailBody_(report, {
    intro: '日報が送信されました。',
    footer: '',
  }));
}

function sendUserDailyReportMail_(mailTo, report) {
  sendMail_(mailTo, buildUserMailSubject_(report), buildDailyReportMailBody_(report, {
    intro: '日報を受け付けました。以下の内容で登録されています。',
    footer: '修正が必要な場合は管理者へご連絡ください。',
  }));
}

function buildAdminMailSubject_(report) {
  return `[日報][管理者] ${getReportWorkDateDisplay_(report)} ${report.name}`;
}

function buildUserMailSubject_(report) {
  return `[日報受付] ${getReportWorkDateDisplay_(report)} ${report.name}`;
}

function buildDailyReportMailBody_(report, options) {
  const intro = String((options && options.intro) || '').trim();
  const footer = String((options && options.footer) || '').trim();

  const lines = [];
  if (intro) {
    lines.push(intro);
    lines.push('');
  }

  lines.push(`氏名: ${report.name}`);
  lines.push(`作業日: ${getReportWorkDateDisplay_(report)}`);
  lines.push(`日報ID: ${report.reportId}`);
  lines.push(`送信回数: ${Number(report.reportCount || 0)} 回`);
  lines.push(`1日合計作業時間: ${Number(report.dailyTotal || report.cumulativeTotal || 0)} h`);
  lines.push(`最新送信日時: ${formatMailDateTime_(report.submittedAt)}`);
  lines.push('');
  lines.push('【作業報告】');

  (report.rows || []).forEach((row, index) => {
    if (index > 0) {
      lines.push('');
    }

    lines.push(`${index + 1}. 依頼No: ${row.requestNo}`);
    lines.push(`   作業内容: ${row.workContent}`);
    lines.push(`   作業時間: ${row.workHours} h`);
  });

  lines.push('');
  lines.push('【一日の総括】');
  lines.push(report.diary || '記載なし');

  if (footer) {
    lines.push('');
    lines.push(footer);
  }

  return lines.join('\n');
}

function sendMail_(to, subject, body) {
  MailApp.sendEmail({
    to,
    subject,
    body,
  });
}

function normalizeMailRecipients_(value) {
  return String(value || '')
    .split(/[\n,;]/)
    .map((item) => item.trim())
    .filter((item) => item !== '')
    .join(',');
}

function getReportWorkDateDisplay_(report) {
  return String((report && (report.workDateDisplay || report.workDate)) || '').trim();
}

function formatMailDateTime_(value) {
  const date = new Date(value);
  if (isNaN(date.getTime())) {
    return String(value || '');
  }
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'yyyy/MM/dd HH:mm:ss'
  );
}
