namespace ExcelWriter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Models the columns within a container
    /// </summary>
    internal class RowOrColumnsModel
    {
        #region Private Fields

        private RowOrColumnInfo first;
        private RowOrColumnInfo last;
        private bool isRowModel;

        #endregion Private Fields

        #region Construction

        public RowOrColumnsModel(bool isRowModel)
        {
            this.isRowModel = isRowModel;
        }

        #endregion Construction

        #region Public Properties

        /// <summary>
        /// Gets or sets the first <see cref="RowOrColumnInfo"/> in the model
        /// </summary>
        public RowOrColumnInfo First
        {
            get { return this.first; }
        }

        /// <summary>
        /// Gets the last <see cref="RowOrColumnInfo"/> in the model
        /// </summary>
        public RowOrColumnInfo Last
        {
            get { return this.last; }
        }

        /// <summary>
        /// Counts and returns the number of <see cref="RowOrColumnInfo"/>s in the <see cref="RowOrColumnsModel"/>
        /// </summary>
        public int Count()
        {
            int count = 0;
            RowOrColumnInfo info = this.first;

            while (info != null)
            {
                count++;
                info.ExcelIndex = count;
                info = info.Next;
            }

            return count;
        }

        #endregion Public Properties

        #region Internal Methods

        /// <summary>
        /// Appends a <see cref="RowOrColumnInfo"/> to the end of this <see cref="RowOrColumnsModel"/>.
        /// </summary>
        /// <param name="map">A <see cref="ExcelMapCoOrdinate"/> which relates to the <see cref="RowOrColumnInfo"/> being added.</param>
        internal void Add(ExcelMapCoOrdinate map)
        {
            if (this.first == null)
            {
                this.first = RowOrColumnInfo.Create(map, this.isRowModel);
                this.last = this.first;
            }
            else
            {
                RowOrColumnInfo newInfo = RowOrColumnInfo.Create(map, this.isRowModel);
                this.last.Next = newInfo;
                this.last = newInfo;
            }
        }

        /// <summary>
        /// Appends the <see cref="RowOrColumnInfo"/>s within a <see cref="RowOrColumnsModel"/> to the end of this <see cref="RowOrColumnsModel"/>.
        /// </summary>
        /// <param name="model"><see cref="RowOrColumnsModel"/> to be appended</param>
        internal void AppendModel(RowOrColumnsModel model)
        {
            if (this.first == null)
            {
                // Model currently not populated, so set
                this.first = model.First;
                this.last = model.Last;
            }
            else
            {
                // Chain to the end
                RowOrColumnInfo nextFirstInfo = model.First;

                this.last.Next = nextFirstInfo;
                if (nextFirstInfo != null)
                {
                    this.last = model.Last;
                }
            }
        }

        /// <summary>
        /// Merge the supplied <see cref="RowOrColumnsModel"/> into this <see cref="RowOrColumnsModel"/>.
        /// </summary>
        /// <param name="modelToBeMerged">The <see cref="RowOrColumnsModel" /> to be merged into this <see cref="RowOrColumnsModel"/></param>
        internal void MergeModel(RowOrColumnsModel modelToBeMerged)
        {
            if (this.first == null)
            {
                // Model currently not populated, so set
                this.first = modelToBeMerged.First;
                this.last = modelToBeMerged.Last;
            }
            else
            {
                // Traverse this linked list.
                RowOrColumnInfo sourceRowOrColInfo = modelToBeMerged.First;

                RowOrColumnInfo previousTargetRowOrColInfo = null;
                RowOrColumnInfo currentTargetRowOrColInfo = this.first;

                while (sourceRowOrColInfo != null)
                {
                    if (!sourceRowOrColInfo.Hidden && !currentTargetRowOrColInfo.Hidden)
                    {
                        // Allocate the visible source column to the visible target columns
                        currentTargetRowOrColInfo = AllocateVisibleSourceToVisibleTargets(sourceRowOrColInfo, currentTargetRowOrColInfo);
                    }
                    else if (sourceRowOrColInfo.Hidden && !currentTargetRowOrColInfo.Hidden)
                    {
                        // Allocate the hidden source column before the current visible target column
                        RowOrColumnInfo insertedColumn = RowOrColumnInfo.Create(this.isRowModel);
                        insertedColumn.Hidden = sourceRowOrColInfo.Hidden;
                        insertedColumn.HeightOrWidth = sourceRowOrColInfo.HeightOrWidth;
                        insertedColumn.AddMaps(sourceRowOrColInfo);
                        insertedColumn.AddMaps(currentTargetRowOrColInfo);

                        insertedColumn.Next = currentTargetRowOrColInfo;

                        if (previousTargetRowOrColInfo == null)
                        {
                            // The new hidden target column is first column in this model
                            this.first = insertedColumn;
                        }
                        else
                        {
                            previousTargetRowOrColInfo.Next = insertedColumn;
                        }

                        // Move back to newly inserted column, so next current remains unchanged
                        currentTargetRowOrColInfo = insertedColumn;
                    }
                    else if (sourceRowOrColInfo.Hidden && currentTargetRowOrColInfo.Hidden)
                    {
                        // Allocate the source to the current target (creating new target row/columns for splits etc).
                        currentTargetRowOrColInfo = AllocateTargetToSource(previousTargetRowOrColInfo, sourceRowOrColInfo);
                    }
                    else if (!sourceRowOrColInfo.Hidden && currentTargetRowOrColInfo.Hidden)
                    {
                        currentTargetRowOrColInfo = AllocateVisibleSourceToTargets(sourceRowOrColInfo, currentTargetRowOrColInfo, sourceRowOrColInfo.HeightOrWidth);
                    }

                    // Source column allocation complete, move on to next source column.
                    sourceRowOrColInfo = sourceRowOrColInfo.Next;

                    // If there is a next source column, and no next target column, then create
                    // a null width target column with the same visibility as the source to accept the next source.
                    if (sourceRowOrColInfo != null && currentTargetRowOrColInfo.Next == null)
                    {
                        currentTargetRowOrColInfo.Next = RowOrColumnInfo.Create(this.isRowModel);
                        currentTargetRowOrColInfo.Next.Hidden = sourceRowOrColInfo.Hidden;

                        // Update the last column if the inserted column is now the last column.
                        if (currentTargetRowOrColInfo == this.last)
                        {
                            this.last = currentTargetRowOrColInfo.Next;
                        }
                    }

                    // Move on to next target column
                    previousTargetRowOrColInfo = currentTargetRowOrColInfo;
                    currentTargetRowOrColInfo = currentTargetRowOrColInfo.Next;
                }
            }
        }

        #endregion Internal Methods

        #region Private Helpers

        /// <summary>
        /// Allocates target <see cref="RowOrColumnInfo"/> to the source <see cref="RowOrColumnInfo"/> until the source row or column is covered.
        /// </summary>
        /// <param name="lastAllocatedTarget">The target row or column where the source row or column is to be allocated</param>
        /// <param name="source">The source row or column to be allocated</param>
        /// <returns>The last target that was allocated (target may have been split to accomodate the source)</returns>
        private RowOrColumnInfo AllocateTargetToSource(RowOrColumnInfo lastAllocatedTarget, RowOrColumnInfo source)
        {
            // Determine the next target column to be allocated to the supplied source
            RowOrColumnInfo lastTarget = lastAllocatedTarget == null ? this.first : lastAllocatedTarget.Next;;

            bool allocateToTargetWidth = lastTarget.HeightOrWidth.HasValue && !source.HeightOrWidth.HasValue;
            bool allocateToSourceWidth = source.HeightOrWidth.HasValue && !lastTarget.HeightOrWidth.HasValue;
            double? remainingSourceWidth = source.HeightOrWidth - lastTarget.HeightOrWidth;

            // Next allocate source to next targets, until all has been allocated
            if (allocateToTargetWidth)
            {
                // Target column has width, but source does not, assign source column info to target.
                lastTarget.MoveMaps(source);
            }
            else if (allocateToSourceWidth)
            {
                // Target column has no specified width, but the source does, assign source column info to target and set width of target to source.
                lastTarget.HeightOrWidth = source.HeightOrWidth;
                lastTarget.MoveMaps(source);
            }
            else if (remainingSourceWidth.HasValue)
            {
                if (remainingSourceWidth.Value == 0)
                {
                    // Source and target column widths are the same, assign source to target
                    lastTarget.MoveMaps(source);
                }
                else if (remainingSourceWidth.Value < 0)
                {
                    // Target column is wider than the source (both have values)

                    // Create a new column for the remaining source width, assign target column info to this column.
                    RowOrColumnInfo insertedTargetColumnInfo = RowOrColumnInfo.Create(this.isRowModel);
                    insertedTargetColumnInfo.HeightOrWidth = -remainingSourceWidth;
                    insertedTargetColumnInfo.Hidden = source.Hidden;
                    insertedTargetColumnInfo.AddMaps(lastTarget);

                    // Truncate the target to the width of the source
                    lastTarget.HeightOrWidth = source.HeightOrWidth;

                    // Assign source column info to target.
                    lastTarget.MoveMaps(source);

                    // Link new column into chain after original target
                    insertedTargetColumnInfo.Next = lastTarget.Next;
                    lastTarget.Next = insertedTargetColumnInfo;

                    // Update the last column if the inserted column is now the last column.
                    if (lastTarget == this.last)
                    {
                        this.last = insertedTargetColumnInfo;
                    }
                }
                else
                {
                    // Source column is wider than the target (both have values)
                    lastTarget = AllocateVisibleSourceToVisibleTargets(source, lastTarget);
                }
            }
            else
            {
                lastTarget.MoveMaps(source);
            }

            // Return the last target column that was allocated
            return lastTarget;
        }

        private RowOrColumnInfo AllocateVisibleSourceToVisibleTargets(RowOrColumnInfo source, RowOrColumnInfo target)
        {
            RowOrColumnInfo currentTarget = target;
            RowOrColumnInfo lastTarget = currentTarget;

            if (currentTarget.HeightOrWidth.HasValue)
            {
                if (source.HeightOrWidth.HasValue)
                {
                    // The amount the source width is over the target width
                    double remainingSourceHeightOrWidth = source.HeightOrWidth.Value - lastTarget.HeightOrWidth.Value;

                    if (remainingSourceHeightOrWidth < 0)
                    {
                        // Target column is wider than the source

                        // Create a new row/column for the remaining source width/height, assign target info to this column.
                        RowOrColumnInfo insertedTarget = RowOrColumnInfo.Create(this.isRowModel);
                        insertedTarget.HeightOrWidth = -remainingSourceHeightOrWidth;
                        insertedTarget.AddMaps(lastTarget);

                        // Truncate the target to the height/width of the source
                        lastTarget.HeightOrWidth = source.HeightOrWidth;

                        // Assign source info to truncated target.
                        lastTarget.MoveMaps(source);

                        // Link new row/column into chain after original target
                        insertedTarget.Next = lastTarget.Next;
                        lastTarget.Next = insertedTarget;

                        // Update the last row/column if the inserted row/column is now the last.
                        if (lastTarget == this.last)
                        {
                            this.last = insertedTarget;
                        }
                    }
                    else
                    {
                        // Source is the same size, or wider, that the target.
                        if (source.Hidden)
                        {
                            lastTarget = AllocateHiddenSourceToTargets(source, lastTarget, remainingSourceHeightOrWidth);
                        }
                        else
                        {
                            lastTarget = AllocateVisibleSourceToTargets(source, lastTarget, remainingSourceHeightOrWidth);
                        }
                    }
                }
                else
                {
                    // Target row/column has height/width, but source does not, assign source info to target.
                    currentTarget.MoveMaps(source);
                }
            }
            else
            {
                // Target row/column has no specified height/width, assign source info to target and set height/width of target to source.
                currentTarget.HeightOrWidth = source.HeightOrWidth;
                currentTarget.MoveMaps(source);
            }
            return lastTarget;
        }

        private RowOrColumnInfo AllocateVisibleSourceToTargets(RowOrColumnInfo source, RowOrColumnInfo target, double? remainingSourceHeightOrWidth)
        {
            // First, assign source info to target.
            if (!(remainingSourceHeightOrWidth.HasValue && remainingSourceHeightOrWidth.Value > 0))
            {
                // Full allocation from source to target...
                target.MoveMaps(source);
            }
            else
            {
                // Partial allocation from source to target.. Leave maps on source.
                target.AddMaps(source);
            }

            // Next allocate source to next targets, until all has been allocated
            while (remainingSourceHeightOrWidth.HasValue && remainingSourceHeightOrWidth.Value > 0)
            {
                // While the target is smaller than the source, allocate new columns 
                // for the target to accept the source column information.
                RowOrColumnInfo nextTarget = target.Next;
                if (nextTarget == null)
                {
                    // No 'Next' target.
                    // Create and link target column for the remaining width, assign source column info to new target.
                    nextTarget = RowOrColumnInfo.Create(this.isRowModel);
                    nextTarget.HeightOrWidth = remainingSourceHeightOrWidth;
                    nextTarget.Hidden = source.Hidden;
                    nextTarget.MoveMaps(source);
                    target.Next = nextTarget;

                    // Update the last row/column if the inserted row/column is now the last.
                    if (target == this.last)
                    {
                        this.last = nextTarget;
                    }

                    // Set remains to 0 so no more allocating.
                    target = nextTarget;
                    remainingSourceHeightOrWidth = 0;
                }
                else
                {
                    if (nextTarget.Hidden)
                    {
                        // Allocate + move on to next
                        nextTarget.AddMaps(source);
                        target = nextTarget;
                    }
                    else
                    {
                        // We have a 'Next' target column which is not hidden
                        if (nextTarget.HeightOrWidth.HasValue)
                        {
                            // It has a fixed width
                            double nextTargetHeightOrWidth = nextTarget.HeightOrWidth.Value;

                            if (nextTargetHeightOrWidth == remainingSourceHeightOrWidth)
                            {
                                // Next target column can be fully allocated to the source
                                nextTarget.MoveMaps(source);
                                target = nextTarget;

                                // Set remains to 0 so no more allocating.
                                remainingSourceHeightOrWidth = 0;
                            }
                            else if (nextTargetHeightOrWidth > remainingSourceHeightOrWidth)
                            {
                                // Next visible target row/column is highter/wider than required remaining.
                                // Create a new target for the remaining + insert before the target
                                RowOrColumnInfo insertedTarget = RowOrColumnInfo.Create(this.isRowModel);
                                insertedTarget.HeightOrWidth = remainingSourceHeightOrWidth;

                                // Assign target + source columns info to inserted.
                                insertedTarget.MoveMaps(source);
                                insertedTarget.AddMaps(nextTarget);

                                // Truncate the original target and update with source column info.
                                nextTarget.HeightOrWidth = nextTargetHeightOrWidth - remainingSourceHeightOrWidth;

                                // Link in the newly inserted column before the next
                                insertedTarget.Next = nextTarget;
                                target.Next = insertedTarget;

                                // Set remains to 0 so no more allocating.
                                target = insertedTarget;
                                remainingSourceHeightOrWidth = 0;
                            }
                            else
                            {
                                // Next target column is narrower than required remaining, assign source column info to target,
                                nextTarget.AddMaps(source);
                                target = nextTarget;

                                // Move on to next row/column reducing remaining.
                                remainingSourceHeightOrWidth = remainingSourceHeightOrWidth - nextTargetHeightOrWidth;
                            }
                        }
                        else
                        {
                            // Next target row/column has no specified height/width, assign source to target, and set height/width of source.
                            nextTarget.HeightOrWidth = remainingSourceHeightOrWidth;
                            nextTarget.MoveMaps(source);
                            target = nextTarget;

                            // Set remains to 0 so no more allocating.
                            remainingSourceHeightOrWidth = 0;
                        }
                    }
                }
            }

            return target;
        }

        private RowOrColumnInfo AllocateHiddenSourceToTargets(RowOrColumnInfo source, RowOrColumnInfo target, double? remainingSourceHeightOrWidth)
        {
            // First, assign source info to target.
            target.AddMaps(source);

            // Next allocate source to next targets, until all has been allocated
            while (remainingSourceHeightOrWidth == null || remainingSourceHeightOrWidth.Value > 0)
            {
                // While the target is smaller than the source, allocate new columns 
                // for the target to accept the source column information.
                RowOrColumnInfo nextTarget = target.Next;
                if (nextTarget == null)
                {
                    // No 'Next' target.
                    // Create and link target row/column for the remaining height/width, assign source to new target.
                    nextTarget = RowOrColumnInfo.Create(this.isRowModel);
                    nextTarget.HeightOrWidth = remainingSourceHeightOrWidth;
                    nextTarget.Hidden = source.Hidden;
                    nextTarget.MoveMaps(source);
                    target.Next = nextTarget;

                    // Update the last row/column if the inserted row/column is now the last.
                    if (target == this.last)
                    {
                        this.last = nextTarget;
                    }

                    // Set remains to 0 so no more allocating.
                    target = nextTarget;
                    remainingSourceHeightOrWidth = 0;
                }
                else
                {
                    if (!nextTarget.Hidden)
                    {
                        // Create and link in a new target for the remaining source.
                        RowOrColumnInfo insertedTarget = RowOrColumnInfo.Create(this.isRowModel);
                        insertedTarget.HeightOrWidth = remainingSourceHeightOrWidth;
                        insertedTarget.Hidden = true;
                        insertedTarget.MoveMaps(source);

                        // Link in the newly inserted column before the next
                        insertedTarget.Next = nextTarget;
                        target.Next = insertedTarget;

                        // Set remains to 0 so no more allocating.
                        target = insertedTarget;
                        remainingSourceHeightOrWidth = 0;
                    }
                    else
                    {
                        // We have a 'Next' target column which is hidden
                        if (nextTarget.HeightOrWidth.HasValue)
                        {
                            // It has a fixed height/width
                            double nextTargetHeightOrWidth = nextTarget.HeightOrWidth.Value;

                            if (nextTargetHeightOrWidth == remainingSourceHeightOrWidth)
                            {
                                // Next target row/column can be fully allocated to the source
                                nextTarget.MoveMaps(source);
                                target = nextTarget;

                                // Set remains to 0 so no more allocating.
                                remainingSourceHeightOrWidth = 0;
                            }
                            else if (nextTargetHeightOrWidth > remainingSourceHeightOrWidth)
                            {
                                // Next target row/column is higher/wider than required remaining.
                                // Create a new target for the remaining height/width + insert before the target
                                RowOrColumnInfo insertedTarget = RowOrColumnInfo.Create(this.isRowModel);
                                insertedTarget.HeightOrWidth = remainingSourceHeightOrWidth;
                                insertedTarget.Hidden = nextTarget.Hidden;

                                // Assign target + source info to inserted.
                                insertedTarget.MoveMaps(source);
                                insertedTarget.AddMaps(nextTarget);

                                // Truncate the original target and update with source info.
                                nextTarget.HeightOrWidth = nextTargetHeightOrWidth - remainingSourceHeightOrWidth;

                                // Link in the newly inserted row/column before the next
                                insertedTarget.Next = nextTarget;
                                target.Next = insertedTarget;

                                // Set remains to 0 so no more allocating.
                                target = insertedTarget;
                                remainingSourceHeightOrWidth = 0;
                            }
                            else
                            {
                                // Next target is smaller than required remaining, assign source info to target,
                                nextTarget.AddMaps(source);
                                target = nextTarget;

                                // Move on to next, reducing remaining.
                                remainingSourceHeightOrWidth = remainingSourceHeightOrWidth - nextTargetHeightOrWidth;
                            }
                        }
                        else
                        {
                            // Next target has no specified height/width, assign source info to target, and set height/width of source.
                            nextTarget.HeightOrWidth = remainingSourceHeightOrWidth;
                            nextTarget.MoveMaps(source);

                            // Set remains to 0 so no more allocating.
                            remainingSourceHeightOrWidth = 0;
                        }
                    }
                }
            }

            return target;
        }

        #endregion Private Helpers
    }
}
