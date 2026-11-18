import React, { useState, useMemo, useEffect } from 'react';
import { 
  Search, 
  CheckCircle2, 
  Circle, 
  Users, 
  CreditCard, 
  Ticket,
  AlertCircle,
  Filter,
  RefreshCw,
  Loader2
} from 'lucide-react';

// ⚠️ PASTE YOUR GOOGLE APPS SCRIPT WEB APP URL HERE ⚠️
const GOOGLE_WEB_APP_URL = "https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec";

export default function App() {
  const [participants, setParticipants] = useState([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSyncing, setIsSyncing] = useState(false);
  const [error, setError] = useState(null);
  
  const [searchQuery, setSearchQuery] = useState('');
  const [filter, setFilter] = useState('all'); // 'all', 'checked-in', 'pending', 'unpaid'

  // Fetch data from Google Sheets on load
  const fetchParticipants = async () => {
    setIsLoading(true);
    setError(null);
    try {
      const response = await fetch(GOOGLE_WEB_APP_URL);
      if (!response.ok) throw new Error("Network response was not ok");
      const data = await response.json();
      setParticipants(data);
    } catch (err) {
      console.error("Fetch error:", err);
      setError("Failed to load data from Google Sheets. Ensure you pasted the correct Web App URL.");
      
      // Fallback data for demonstration if URL isn't set
      if (GOOGLE_WEB_APP_URL.includes("YOUR_SCRIPT_ID")) {
        setParticipants([
          { id: 2, firstName: "Demo", lastName: "Attendee", email: "attendee@example.com", ticketCode: "TICKET-SAMPLE-01", numAttendees: 1, totalCost: 0, paymentStatus: true, checkedIn: false }
        ]);
        setError("Please replace GOOGLE_WEB_APP_URL with your actual Google Apps Script URL.");
      }
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    fetchParticipants();
  }, []);

  // Toggle check-in status and sync with Google Sheets
  const toggleCheckIn = async (id, currentStatus) => {
    const newStatus = !currentStatus;
    
    // Optimistic UI update (feels instant to the user)
    setParticipants(prev => prev.map(p => p.id === id ? { ...p, checkedIn: newStatus } : p));
    setIsSyncing(true);

    try {
      // Send update to Google Sheets (POST request)
      const response = await fetch(GOOGLE_WEB_APP_URL, {
        method: 'POST',
        // 'text/plain' helps bypass strict CORS preflight checks in some environments
        headers: { 'Content-Type': 'text/plain;charset=utf-8' }, 
        body: JSON.stringify({ id: id, checkedIn: newStatus })
      });
      
      const result = await response.json();
      if (!result.success) throw new Error("Failed to update sheet");
      
    } catch (err) {
      console.error("Sync error:", err);
      alert("Failed to sync check-in to Google Sheets. Check connection.");
      // Rollback the UI if it failed
      setParticipants(prev => prev.map(p => p.id === id ? { ...p, checkedIn: currentStatus } : p));
    } finally {
      setIsSyncing(false);
    }
  };

  // Derived statistics
  const stats = useMemo(() => {
    let totalExpected = 0;
    let totalCheckedIn = 0;
    let totalUnpaid = 0;

    participants.forEach(p => {
      totalExpected += p.numAttendees;
      if (p.checkedIn) totalCheckedIn += p.numAttendees;
      if (!p.paymentStatus) totalUnpaid++;
    });

    return { totalExpected, totalCheckedIn, totalUnpaid };
  }, [participants]);

  // Filter and search logic
  const filteredParticipants = useMemo(() => {
    return participants.filter(p => {
      const searchLower = searchQuery.toLowerCase();
      const matchesSearch = 
        p.firstName.toLowerCase().includes(searchLower) ||
        p.lastName.toLowerCase().includes(searchLower) ||
        p.ticketCode.toLowerCase().includes(searchLower) ||
        p.email.toLowerCase().includes(searchLower);

      if (!matchesSearch) return false;

      if (filter === 'checked-in') return p.checkedIn;
      if (filter === 'pending') return !p.checkedIn;
      if (filter === 'unpaid') return !p.paymentStatus;
      
      return true;
    });
  }, [participants, searchQuery, filter]);

  if (isLoading) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center text-slate-500">
        <Loader2 className="animate-spin mb-4" size={40} />
        <p className="font-medium">Connecting to Google Sheets...</p>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans pb-20">
      {/* Header */}
      <header className="bg-indigo-600 text-white p-4 shadow-md sticky top-0 z-10">
        <div className="max-w-3xl mx-auto flex items-center justify-between">
          <div>
            <h1 className="text-xl font-bold">Event Check-In</h1>
            <p className="text-indigo-200 text-sm flex items-center gap-2">
              Live Attendance & Verification
              {isSyncing && <Loader2 size={12} className="animate-spin text-indigo-200" />}
            </p>
          </div>
          <div className="flex items-center gap-3">
            <button 
              onClick={fetchParticipants}
              className="p-2 rounded-full hover:bg-indigo-500 transition-colors"
              title="Refresh Data"
            >
              <RefreshCw size={18} className={isSyncing ? "animate-spin" : ""} />
            </button>
            <div className="bg-indigo-500 px-3 py-1 rounded-full text-sm font-medium border border-indigo-400">
              {stats.totalCheckedIn} / {stats.totalExpected} Here
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-3xl mx-auto p-4 space-y-6">
        {error && (
          <div className="bg-red-50 text-red-700 p-4 rounded-xl border border-red-200 text-sm flex gap-3">
            <AlertCircle className="shrink-0" size={20} />
            <p>{error}</p>
          </div>
        )}

        {/* Statistics Cards */}
        <div className="grid grid-cols-3 gap-3">
          <div className="bg-white p-3 rounded-xl shadow-sm border border-slate-200 flex flex-col items-center justify-center text-center">
            <Users className="text-indigo-500 mb-1" size={24} />
            <span className="text-2xl font-bold text-slate-800">{stats.totalExpected}</span>
            <span className="text-xs text-slate-500 font-medium uppercase tracking-wider">Expected</span>
          </div>
          <div className="bg-white p-3 rounded-xl shadow-sm border border-emerald-200 flex flex-col items-center justify-center text-center">
            <CheckCircle2 className="text-emerald-500 mb-1" size={24} />
            <span className="text-2xl font-bold text-emerald-700">{stats.totalCheckedIn}</span>
            <span className="text-xs text-emerald-600 font-medium uppercase tracking-wider">Checked In</span>
          </div>
          <div className={`bg-white p-3 rounded-xl shadow-sm border flex flex-col items-center justify-center text-center ${stats.totalUnpaid > 0 ? 'border-amber-200' : 'border-slate-200'}`}>
            <AlertCircle className={stats.totalUnpaid > 0 ? "text-amber-500 mb-1" : "text-slate-400 mb-1"} size={24} />
            <span className={`text-2xl font-bold ${stats.totalUnpaid > 0 ? 'text-amber-600' : 'text-slate-600'}`}>{stats.totalUnpaid}</span>
            <span className={`text-xs font-medium uppercase tracking-wider ${stats.totalUnpaid > 0 ? 'text-amber-600' : 'text-slate-500'}`}>Unpaid</span>
          </div>
        </div>

        {/* Search and Filters */}
        <div className="space-y-3">
          <div className="relative">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={20} />
            <input 
              type="text" 
              placeholder="Search by name, email, or ticket code (WD26-...)"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full pl-10 pr-4 py-3 bg-white border border-slate-300 rounded-xl shadow-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 transition-all"
            />
          </div>

          <div className="flex gap-2 overflow-x-auto pb-2 scrollbar-hide">
            <button 
              onClick={() => setFilter('all')}
              className={`whitespace-nowrap px-4 py-2 rounded-full text-sm font-medium transition-colors ${filter === 'all' ? 'bg-indigo-600 text-white' : 'bg-white border border-slate-300 text-slate-600 hover:bg-slate-50'}`}
            >
              All Participants
            </button>
            <button 
              onClick={() => setFilter('pending')}
              className={`whitespace-nowrap px-4 py-2 rounded-full text-sm font-medium transition-colors ${filter === 'pending' ? 'bg-indigo-600 text-white' : 'bg-white border border-slate-300 text-slate-600 hover:bg-slate-50'}`}
            >
              Pending Arrival
            </button>
            <button 
              onClick={() => setFilter('checked-in')}
              className={`whitespace-nowrap px-4 py-2 rounded-full text-sm font-medium transition-colors ${filter === 'checked-in' ? 'bg-emerald-600 text-white' : 'bg-white border border-slate-300 text-slate-600 hover:bg-slate-50'}`}
            >
              Checked In
            </button>
            <button 
              onClick={() => setFilter('unpaid')}
              className={`whitespace-nowrap px-4 py-2 rounded-full text-sm font-medium transition-colors ${filter === 'unpaid' ? 'bg-amber-500 text-white' : 'bg-white border border-slate-300 text-slate-600 hover:bg-slate-50'}`}
            >
              Unpaid Tickets
            </button>
          </div>
        </div>

        {/* Participant List */}
        <div className="space-y-3">
          {filteredParticipants.length === 0 ? (
            <div className="text-center py-10 bg-white rounded-xl border border-slate-200">
              <Filter className="mx-auto h-12 w-12 text-slate-300 mb-3" />
              <p className="text-slate-500 font-medium">No participants found.</p>
            </div>
          ) : (
            filteredParticipants.map(participant => (
              <div 
                key={participant.id} 
                className={`bg-white rounded-xl p-4 shadow-sm border transition-all ${participant.checkedIn ? 'border-emerald-500 ring-1 ring-emerald-500' : 'border-slate-200'}`}
              >
                <div className="flex justify-between items-start mb-3">
                  <div>
                    <h3 className="text-lg font-bold text-slate-900 leading-tight">
                      {participant.firstName} {participant.lastName}
                    </h3>
                    <p className="text-slate-500 text-sm truncate max-w-[200px] sm:max-w-xs">{participant.email}</p>
                  </div>
                  <div className="text-right flex flex-col items-end">
                    <span className="inline-flex items-center gap-1 bg-slate-100 px-2 py-1 rounded text-xs font-mono font-semibold text-slate-700 border border-slate-200">
                      <Ticket size={12} />
                      {participant.ticketCode}
                    </span>
                  </div>
                </div>

                <div className="flex items-center gap-4 mb-4 text-sm">
                  <div className="flex items-center gap-1.5 text-slate-700">
                    <Users size={16} className="text-indigo-500" />
                    <span className="font-medium">{participant.numAttendees} {participant.numAttendees === 1 ? 'Person' : 'People'}</span>
                  </div>
                  <div className="flex items-center gap-1.5">
                    <CreditCard size={16} className={participant.paymentStatus ? "text-emerald-500" : "text-amber-500"} />
                    {participant.paymentStatus ? (
                      <span className="text-emerald-700 font-medium bg-emerald-50 px-2 py-0.5 rounded-md">Paid (€{participant.totalCost})</span>
                    ) : (
                      <span className="text-amber-700 font-bold bg-amber-50 px-2 py-0.5 rounded-md border border-amber-200">Due: €{participant.totalCost}</span>
                    )}
                  </div>
                </div>

                <button
                  onClick={() => toggleCheckIn(participant.id, participant.checkedIn)}
                  disabled={isSyncing}
                  className={`w-full py-3 px-4 rounded-lg font-bold flex items-center justify-center gap-2 transition-all active:scale-[0.98] disabled:opacity-70 ${
                    participant.checkedIn 
                      ? 'bg-emerald-50 text-emerald-700 border border-emerald-200 hover:bg-emerald-100' 
                      : 'bg-indigo-600 text-white shadow-md hover:bg-indigo-700'
                  }`}
                >
                  {participant.checkedIn ? (
                    <>
                      <CheckCircle2 size={20} />
                      Checked In — Tap to Undo
                    </>
                  ) : (
                    <>
                      <Circle size={20} />
                      Check In Arrival
                    </>
                  )}
                </button>
              </div>
            ))
          )}
        </div>
      </main>
    </div>
  );
}
