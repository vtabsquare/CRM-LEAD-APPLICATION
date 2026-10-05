import React, { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { ArrowLeft, Save, AlertCircle } from 'lucide-react';
import api from '../api';

const HomeCareCreate = () => {
    const navigate = useNavigate();
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState('');
    const [success, setSuccess] = useState('');

    const [formData, setFormData] = useState({
        patient_name: '',
        gender: '',
        age: '',
        pain_point: '',
        location: '',
        service_started_on: new Date().toISOString().split('T')[0],
        service_type: 'Home Care',
        home_care_revenue: '',
        additional_nursing_charges: '0',
        discount: '0',
        shift: 'Regular',
        active_inactive: 'ACTIVE',
        service_stopped_on: '',
    });

    const handleChange = (e) => {
        const { name, value } = e.target;
        setFormData(prev => ({
            ...prev,
            [name]: value
        }));
    };

    const handleSubmit = async (e) => {
        e.preventDefault();
        setError('');
        setSuccess('');

        // Validation
        if (!formData.patient_name.trim()) {
            setError('Patient name is required');
            return;
        }
        if (!formData.home_care_revenue || parseFloat(formData.home_care_revenue) <= 0) {
            setError('Home care revenue must be greater than 0');
            return;
        }

        try {
            setLoading(true);

            // Use centralized API client which automatically attaches JWT token
            const response = await api.post('/api/homecare/clients', formData);
            const data = response.data;

            if (data.status === 'success') {
                setSuccess('Home care client added successfully!');
                setTimeout(() => {
                    navigate('/homecare/clients');
                }, 2000);
            } else {
                setError(data.detail || 'Failed to add client');
            }
        } catch (err) {
            console.error('Error adding client:', err);
            // Check if it's an Axios error with a response from the backend
            if (err.response && err.response.data && err.response.data.detail) {
                setError(err.response.data.detail);
            } else {
                setError('Failed to add client. Please try again.');
            }
        } finally {
            setLoading(false);
        }
    };


    return (
        <div className="p-6">
            {/* Header */}
            <div className="mb-6">
                <button
                    onClick={() => navigate('/homecare/clients')}
                    className="flex items-center gap-2 text-blue-600 hover:text-blue-800 mb-4"
                >
                    <ArrowLeft size={20} />
                    Back to Clients
                </button>

                <h1 className="text-2xl font-bold text-gray-800">Add New Home Care Client</h1>
                <p className="text-gray-600 mt-1">Enter patient details for home care service</p>
            </div>

            {/* Error/Success Messages */}
            {error && (
                <div className="mb-4 p-4 bg-red-100 border border-red-400 text-red-700 rounded-lg flex items-center gap-2">
                    <AlertCircle size={20} />
                    {error}
                </div>
            )}

            {success && (
                <div className="mb-4 p-4 bg-green-100 border border-green-400 text-green-700 rounded-lg">
                    {success}
                </div>
            )}

            {/* Form */}
            <form onSubmit={handleSubmit} className="bg-white rounded-lg shadow p-6">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    {/* Patient Name */}
                    <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                            Patient Name <span className="text-red-500">*</span>
                        </label>
                        <input
                            type="text"
                            name="patient_name"
                            value={formData.patient_name}
                            onChange={handleChange}
                            className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                            placeholder="Enter patient name"
                            required
                        />
                    </div>

                    {/* Gender */}
                    <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                            Gender
                        </label>
                        <select
                            name="gender"
                            value={formData.gender}
                            onChange={handleChange}
                            className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                        >
                            <option value="">Select Gender</option>
                            <option value="Male">Male</option>
                            <option value="Female">Female</option>
                            <option value="Other">Other</option>
                        </select>
                    </div>

                    {/* Age */}
                    <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                            Age
                        </label>
                        <input
                            type="number"
                            name="age"
                            value={formData.age}
                            onChange={handleChange}
                            className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                            min="0"
                            max="150"
                        />
                    </div>

                    {/* Location */}
                    <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                            Location
                        </label>
                        <input
                            type="text"
                            name="location"
                            value={formData.location}
                            onChange={handleChange}
                            className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                        />
                    </div>

                    {/* Pain Point */}
                    <div className="md:col-span-2">
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                            Pain Point / Diagnosis
                        </label>
                        <textarea
                            name="pain_point"
                            value={formData.pain_point}
                            onChange={handleChange}
                            rows="3"
                            className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                        />
                    </div>

                    {/* Service Started On */}
                    <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                            Service Started On <span className="text-red-500">*</span>
                        </label>
                        <input
                            type="date"
                            name="service_started_on"
                            value={formData.service_started_on}
                            onChange={handleChange}
                            className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                            required
                        />
                    </div>

                    {/* Shift */}
                    <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                            Shift
                        </label>
                        <select
                            name="shift"
                            value={formData.shift}
                            onChange={handleChange}
                            className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                        >
                            <option value="Regular">Regular</option>
                            <option value="Day">Day</option>
                            <option value="Night">Night</option>
                            <option value="24/7">24/7</option>
                        </select>
                    </div>

                    {/* Home Care Revenue */}
                    <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                            Home Care Revenue (₹) <span className="text-red-500">*</span>
                        </label>
                        <input
                            type="number"
                            name="home_care_revenue"
                            value={formData.home_care_revenue}
                            onChange={handleChange}
                            className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                            min="0"
                            step="0.01"
                            required
                        />
                    </div>

                    {/* Additional Nursing Charges */}
                    <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                            Additional Nursing Charges (₹)
                        </label>
                        <input
                            type="number"
                            name="additional_nursing_charges"
                            value={formData.additional_nursing_charges}
                            onChange={handleChange}
                            className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                            min="0"
                            step="0.01"
                        />
                    </div>

                    {/* Discount */}
                    <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                            Discount (₹)
                        </label>
                        <input
                            type="number"
                            name="discount"
                            value={formData.discount}
                            onChange={handleChange}
                            className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                            min="0"
                            step="0.01"
                        />
                    </div>

                    {/* Total Revenue (Calculated) */}
                    <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                            Total Revenue (₹)
                        </label>
                        <div className="w-full px-4 py-2 bg-gray-100 border border-gray-300 rounded-lg text-gray-700 font-semibold">
                            ₹{(
                                (parseFloat(formData.home_care_revenue) || 0) +
                                (parseFloat(formData.additional_nursing_charges) || 0) -
                                (parseFloat(formData.discount) || 0)
                            ).toLocaleString()}
                        </div>
                    </div>
                </div>

                {/* Action Buttons */}
                <div className="mt-6 flex gap-4">
                    <button
                        type="submit"
                        disabled={loading}
                        className="px-6 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 disabled:bg-gray-400 flex items-center gap-2"
                    >
                        <Save size={20} />
                        {loading ? 'Saving...' : 'Save Client'}
                    </button>
                    <button
                        type="button"
                        onClick={() => navigate('/homecare/clients')}
                        className="px-6 py-2 bg-gray-200 text-gray-700 rounded-lg hover:bg-gray-300"
                    >
                        Cancel
                    </button>
                </div>
            </form>
        </div>
    );
};

export default HomeCareCreate;
