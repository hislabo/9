 // Top navbar: live date/time and button handlers
    function _updateTopDateTime(){
      const now = new Date();
      const dateEl = document.getElementById('currentDate');
      const timeEl = document.getElementById('currentClock');
      if(!dateEl || !timeEl) return;
      const dateStr = now.toLocaleDateString('vi-VN', { weekday: 'long', year: 'numeric', month: 'short', day: 'numeric' });
      const timeStr = now.toLocaleTimeString('vi-VN');
      dateEl.textContent = dateStr;
      timeEl.textContent = timeStr;
    }
    document.addEventListener('DOMContentLoaded', ()=>{
      _updateTopDateTime();
      const bsModalPatient = new bootstrap.Modal(document.getElementById('patientModal'));
      document.getElementById('topAddBtn')?.addEventListener('click', ()=> bsModalPatient.show());
      document.getElementById('topDeleteBtn')?.addEventListener('click', ()=> {
        if (!selectedPatientIds.size) return alert('Vui lòng chọn ít nhất một bệnh nhân để xóa.');
        if (!confirm(`Xác nhận xóa ${selectedPatientIds.size} bệnh nhân đã chọn?`)) return;
        deleteSelectedPatients();
      });
      document.getElementById('topRefreshBtn')?.addEventListener('click', ()=> location.reload());
      document.getElementById('topPrintBtn')?.addEventListener('click', ()=> window.print());
      document.getElementById('topSaveSettingsBtn')?.addEventListener('click', ()=> alert('Cài đặt đã được lưu.'));
      document.getElementById('topCalendarBtn')?.addEventListener('click', ()=> {
        reminderStatusFilter = 'all';
        reminderSearchFilter = '';
        const searchInput = document.getElementById('reminderSearchInput');
        const statusSelect = document.getElementById('reminderStatusSelect');
        if (searchInput) searchInput.value = '';
        if (statusSelect) statusSelect.value = 'all';
        showSection('reminderSection');
        renderReminders();
      });
      document.getElementById('topPaymentBtn')?.addEventListener('click', ()=> alert('Mở chức năng thanh toán (tạm).'));
      document.getElementById('topLogoutBtn')?.addEventListener('click', ()=> document.getElementById('logoutButton')?.click());
    });
